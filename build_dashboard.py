# -*- coding: utf-8 -*-
"""
Citizen Property Tax & RS Assessment Survey (v25) — Dashboard Builder
=====================================================================
Research Solutions (M&A Research Solutions LLC)

Reads the daily SurveyCTO WIDE export + the sampling frame (rs_prefill.xlsx)
+ the XLSForm (citizen_v25.xlsx, for value labels), computes every aggregate
the dashboard needs, and injects the JSON into dashboard_template.html to
produce index.html (the file served on GitHub Pages).

Daily usage:
    1. Drop the fresh export over
       "Citizen Property Tax & RS Assessment Survey (v25)_WIDE.csv"
    2. Run:  python build_dashboard.py     (or double-click update_dashboard.bat)
"""

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).parent
CSV_PATH = HERE / "Citizen Property Tax & RS Assessment Survey (v25)_WIDE.csv"
PREFILL_PATH = HERE / "rs_prefill.xlsx"
XLSFORM_PATH = HERE / "citizen_v25.xlsx"
TEMPLATE_PATH = HERE / "dashboard_template.html"
OUT_PATH = HERE / "index.html"

MISSING_CODES = {97, 98, 99, 666, 777, 888, 999}   # excluded from scale means


# ----------------------------------------------------------------------
#  Load
# ----------------------------------------------------------------------
def load_labels():
    """Map question name -> {code: english label} from the XLSForm."""
    sv = pd.read_excel(XLSFORM_PATH, sheet_name="survey")
    ch = pd.read_excel(XLSFORM_PATH, sheet_name="choices")
    lists = {}
    for _, r in ch.iterrows():
        ln = str(r["list_name"]).strip()
        if ln in ("", "nan"):
            continue
        try:
            val = int(float(r["value"]))
        except (ValueError, TypeError):
            val = str(r["value"]).strip()
        lab = str(r.get("label: eng", "")).strip()
        lab = re.sub(r"<[^>]+>", "", lab).strip()
        lists.setdefault(ln, {})[val] = lab
    qmap = {}
    for _, r in sv.iterrows():
        t = str(r["type"]).strip()
        m = re.match(r"select_(one|multiple)\s+(\S+)", t)
        if m:
            qmap[str(r["name"]).strip()] = lists.get(m.group(2), {})
    return qmap, lists


def load_data():
    df = pd.read_csv(CSV_PATH, dtype=str, low_memory=False)
    df.columns = [c.strip() for c in df.columns]
    pf = pd.read_excel(PREFILL_PATH, sheet_name=0)
    return df, pf


# ----------------------------------------------------------------------
#  Helpers
# ----------------------------------------------------------------------
def numcol(s):
    return pd.to_numeric(s, errors="coerce")


def dist(frame, col, labels, order_by_choices=True, drop_missing_codes=False):
    """[{label, value}] counts for a select_one column, in choice order."""
    if col not in frame:
        return []
    v = numcol(frame[col]).dropna().astype(int)
    if drop_missing_codes:
        v = v[~v.isin(MISSING_CODES)]
    counts = v.value_counts()
    lmap = labels.get(col, {})
    out = []
    keys = list(lmap.keys()) if order_by_choices else list(counts.index)
    for k in keys:
        n = int(counts.get(k, 0))
        if n:
            out.append({"label": str(lmap.get(k, k)), "value": n})
    # any codes not in the choice list
    for k in counts.index:
        if k not in lmap:
            out.append({"label": str(k), "value": int(counts[k])})
    return out


def multi(frame, base, labels):
    """[{label, value}] for a select_multiple's dummy columns, desc by count."""
    lmap = labels.get(base, {})
    out = []
    for code, lab in lmap.items():
        col = f"{base}_{code}"
        if col in frame:
            n = int((numcol(frame[col]) == 1).sum())
            if n:
                out.append({"label": lab, "value": n})
    out.sort(key=lambda x: -x["value"])
    return out


def mean_scale(frame, col, decimals=1):
    v = numcol(frame.get(col, pd.Series(dtype=float))).dropna()
    v = v[~v.isin(MISSING_CODES)]
    return round(float(v.mean()), decimals) if len(v) else None


def hist_int(frame, col, lo, hi):
    """Histogram of an integer scale lo..hi (missing codes dropped)."""
    v = numcol(frame.get(col, pd.Series(dtype=float))).dropna()
    v = v[~v.isin(MISSING_CODES)]
    v = v[(v >= lo) & (v <= hi)].astype(int)
    return [{"label": str(i), "value": int((v == i).sum())} for i in range(lo, hi + 1)]


def bins(series, edges, fmt):
    v = numcol(series).dropna()
    out = []
    for i in range(len(edges) - 1):
        a, b = edges[i], edges[i + 1]
        n = int(((v >= a) & (v < b)).sum())
        out.append({"label": fmt(a, b), "value": n})
    return out


def pct_yes(frame, col, yes=1, valid=None):
    v = numcol(frame.get(col, pd.Series(dtype=float))).dropna()
    if valid is not None:
        v = v[v.isin(valid)]
    return round(100 * float((v == yes).mean()), 0) if len(v) else None


def top2_agree_pct(frame, col, top=(3, 4)):
    v = numcol(frame.get(col, pd.Series(dtype=float))).dropna()
    v = v[~v.isin(MISSING_CODES)]
    return round(100 * float(v.isin(top).mean())) if len(v) else None


# ----------------------------------------------------------------------
#  Build
# ----------------------------------------------------------------------
def build():
    labels, lists = load_labels()
    df, pf = load_data()

    # field date = actual interview date (starttime), not server sync date
    st = pd.to_datetime(df["starttime"], errors="coerce", format="mixed")
    df["_fdate"] = st.dt.date.astype(str)

    df["_status"] = numcol(df["survey_status"])
    comp = df[df["_status"] == 1].copy()              # completed interviews
    resp = df[numcol(df["s0_consent_begin"]) == 1].copy()   # consented → asked the questionnaire

    # ---- sampling frame ----
    samp = pf[pf["sampled"] == 1].copy()
    total_target = int(len(samp))
    treat_target = samp["treat"].value_counts().to_dict()
    frame_loc = (
        samp.groupby("locality_name")
        .agg(target=("resp_id", "size"), circle=("circle_name", lambda s: s.mode().iat[0]))
        .reset_index()
    )

    # completed joined to frame (for circle / treat cross-checks)
    comp["_rid"] = numcol(comp["resp_id"])
    pf_idx = pf.set_index("resp_id")
    comp["_circle"] = comp["_rid"].map(pf_idx["circle_name"])

    # ---- meta ----
    dates = sorted(d for d in comp["_fdate"].dropna().unique() if d and d != "NaT")
    dur_min = (numcol(comp["duration"]) / 60).dropna()
    dur_min = dur_min[dur_min > 5]  # guard against timer glitches
    n_complete = int(len(comp))
    n_sub = int(len(df))

    slabs = lists.get("status_survey", {})
    sc = numcol(df["survey_status"]).dropna().astype(int).value_counts()
    status_dist = [{"label": slabs.get(k, str(k)), "value": int(v)} for k, v in sc.items()]
    status_dist.sort(key=lambda x: -x["value"])

    refused = int(numcol(df["survey_status"]).isin([3, 4, 5, 6]).sum())

    treat_done = comp["treatment"].value_counts().to_dict()

    daily = (
        comp.groupby("_fdate").size().reset_index(name="count")
        .rename(columns={"_fdate": "date"}).sort_values("date")
    )
    daily = [{"date": r["date"], "count": int(r["count"])} for _, r in daily.iterrows()]

    # ---- WTP demand curve ----
    w = comp.copy()
    w["_p"] = numcol(w["s6_1_price_offered"])
    w["_b"] = numcol(w["s6_1_buy_own_estimate"])
    w = w[w["_b"].isin([0, 1])]
    curve = []
    for p, g in w.groupby("_p"):
        curve.append({"price": int(p), "n": int(len(g)), "pct": round(100 * float(g["_b"].mean()))})
    curve.sort(key=lambda x: x["price"])

    # ---- treatment-arm comparison (RS attitudes, % top-2 box) ----
    arm_rows = {"Control": comp[comp["treatment"] == "Control"],
                "T1": comp[comp["treatment"] == "T1"],
                "T2": comp[comp["treatment"] == "T2"]}
    byarm_metrics = [
        ("RS system reliable (agree)", lambda g: top2_agree_pct(g, "s5_3_rs_reliable")),
        ("RS treats all fairly (agree)", lambda g: top2_agree_pct(g, "s5_4_rs_fair_post")),
        ("Errors correctable (easy)", lambda g: top2_agree_pct(g, "s5_5_rs_correctable")),
        ("Would buy own estimate", lambda g: pct_yes(g, "s6_1_buy_own_estimate", valid=[0, 1])),
    ]
    byarm = {"labels": [m[0] for m in byarm_metrics]}
    for arm, g in arm_rows.items():
        byarm[arm] = [m[1](g) if len(g) else None for m in byarm_metrics]

    # ---- MC recall (treatment script comprehension) ----
    mc = []
    for col, lab in [
        ("s4_mc0_recall_updatefreq_correct", "Legal update frequency"),
        ("s4_mc1_recall_rs_idea_correct", "What RS valuation uses"),
        ("s4_mc2_whose_house_correct", "Whose house the example was"),
        ("s4_mc3_two_sided_correct", "Estimate can also be smaller"),
    ]:
        v = numcol(comp.get(col, pd.Series(dtype=float))).dropna()
        if len(v):
            mc.append({"label": lab, "value": round(100 * float((v == 1).mean()))})

    # ---- documents ----
    docs = []
    for col, lab in [
        ("s3_16_doc_possess_a", "Registry / sale deed"),
        ("s3_16_doc_possess_b", "Architectural drawings"),
        ("s3_16_doc_possess_c", "Latest PT-1 / tax receipt"),
    ]:
        p = pct_yes(resp, col, valid=[0, 1])
        if p is not None:
            docs.append({"label": lab, "value": int(p)})

    # ---- assets & observations (yn12: 1 yes, 2 no) ----
    assets = []
    for col, lab in [
        ("s12_11_obs_ac", "Air conditioner"),
        ("s12_10_obs_air_cooler", "Air cooler"),
        ("s12_12_obs_motorcycle", "Motorcycle / scooter"),
        ("s12_13_obs_car", "Car or truck"),
    ]:
        p = pct_yes(resp, col, valid=[1, 2])
        if p is not None:
            assets.append({"label": lab, "value": int(p)})
    obs_avgs = []
    for col, lab in [
        ("s12_2_obs_wealthy", "Wealthy"),
        ("s12_3_obs_religious", "Religious"),
        ("s12_4_obs_influential", "Influential"),
        ("s12_5_obs_educated", "Educated"),
    ]:
        m = mean_scale(resp, col)
        if m is not None:
            obs_avgs.append({"label": lab, "value": m})

    # ---- enumerators ----
    en = comp["enum_label"].fillna("Unknown").value_counts()
    enums = [{"label": k, "value": int(v)} for k, v in en.items()][:14]

    # ---- locality table ----
    done_by_loc = comp.groupby("locality_name").size().to_dict()
    loc_rows = []
    for _, r in frame_loc.iterrows():
        done = int(done_by_loc.get(r["locality_name"], 0))
        tgt = int(r["target"])
        pct = round(100 * done / tgt) if tgt else 0
        status = "Completed" if tgt and done >= tgt else ("In Progress" if done else "Not Started")
        loc_rows.append({"loc": r["locality_name"], "circle": str(r["circle"]), "done": done,
                         "target": tgt, "pct": pct, "status": status, "touched": done > 0})
    known = {r["loc"] for r in loc_rows}
    for loc, done in done_by_loc.items():
        if loc not in known and isinstance(loc, str):
            loc_rows.append({"loc": loc, "circle": "—", "done": int(done), "target": 0,
                             "pct": 0, "status": "In Progress", "touched": True})
    loc_rows.sort(key=lambda x: -x["pct"])

    # ---- circle progress (top by completed) ----
    circ_tgt = samp.groupby("circle_name").size().to_dict()
    circ_done = comp.groupby("_circle").size().to_dict()
    circ = [{"label": c, "done": int(d), "target": int(circ_tgt.get(c, 0))}
            for c, d in circ_done.items() if isinstance(c, str)]
    circ.sort(key=lambda x: -x["done"])

    # ---- duration histogram (completed) ----
    dur_hist = bins(dur_min, [0, 40, 50, 60, 75, 90, 10_000],
                    lambda a, b: f"{int(a)}–{int(b)} min" if b < 10_000 else "90+ min")

    now = datetime.now()
    D = {
        "meta": {
            "last_updated": now.strftime("%d %b %Y, %I:%M %p"),
            "last_date": dates[-1] if dates else "",
            "field_days": len(dates),
            "n_submissions": n_sub,
            "n_consented": int(len(resp)),
            "n_complete": n_complete,
            "n_refused": refused,
            "total_target": total_target,
            "pct_complete": round(100 * n_complete / total_target, 1) if total_target else 0,
            "response_rate": round(100 * n_complete / n_sub) if n_sub else 0,
            "median_duration": round(float(dur_min.median())) if len(dur_min) else 0,
            "n_enums": int(comp["enum_label"].nunique()),
            "n_localities_frame": int(len(frame_loc)),
            "n_localities_started": int(sum(1 for r in loc_rows if r["touched"])),
            "n_circles_frame": int(samp["circle_name"].nunique()),
            "n_circles_started": int(len(circ)),
            "t1_done": int(treat_done.get("T1", 0)), "t1_target": int(treat_target.get("T1", 0)),
            "t2_done": int(treat_done.get("T2", 0)), "t2_target": int(treat_target.get("T2", 0)),
            "ctrl_done": int(treat_done.get("Control", 0)), "ctrl_target": int(treat_target.get("Control", 0)),
            "portal_aware_pct": pct_yes(resp, "s1_1_portal_aware", valid=[0, 1]),
            "portal_used_pct": pct_yes(resp, "s1_3_portal_used_12m", valid=[0, 1]),
            "contact6m_pct": pct_yes(resp, "s2_1_et_contact_6m", valid=[0, 1]),
            "officer_trust_pct": top2_agree_pct(resp, "s3_9_officer_trust"),
            "sat_overall_pct": top2_agree_pct(resp, "s3_10_sat_overall"),
            "bribe_asked_pct": pct_yes(resp, "s9_10_bribe_asked", valid=[0, 1]),
            "seen_record_pct": (lambda v: round(100 * float(v.isin([1, 2]).mean())) if len(v) else None)(
                numcol(resp.get("s3_1_seen_record", pd.Series(dtype=float))).dropna().pipe(lambda s: s[~s.isin(MISSING_CODES)])),
            "avg_info_surprising": mean_scale(comp, "s4_3_info_surprising", 0),
            "avg_info_trust": mean_scale(comp, "s4_4_info_trust", 0),
            "avg_rs_accuracy": mean_scale(comp, "s5_2_rs_accuracy_oo10"),
            "avg_prior_conf": mean_scale(comp, "s3_5_prior_confidence"),
            "avg_appeal_int": mean_scale(resp, "s9_1_appeal_intention"),
            "avg_holdout": mean_scale(comp, "s7_holdout_believe"),
            "buy_yes_pct": pct_yes(comp, "s6_1_buy_own_estimate", valid=[0, 1]),
            "avg_liability": (lambda v: int(round(float(v.mean()), -2)) if len(v) else None)(
                numcol(comp.get("initial_liability", pd.Series(dtype=float))).dropna()),
            "rs_reliable_pct": top2_agree_pct(comp, "s5_3_rs_reliable"),
            "rs_fair_pct": top2_agree_pct(comp, "s5_4_rs_fair_post"),
        },
        "daily": daily,
        "status_dist": status_dist,
        "treat_dist": [{"label": k, "value": int(v)} for k, v in treat_done.items()],
        "circle_prog": circ[:12],
        "dur_hist": dur_hist,
        "enums": enums,

        "portal": {
            "aware": dist(resp, "s1_1_portal_aware", labels),
            "known": multi(resp, "s1_2_portal_known", labels),
            "used": dist(resp, "s1_3_portal_used_12m", labels),
            "nonuse": multi(resp, "s1_4_portal_nonuse_why", labels),
            "help": dist(resp, "s1_5_portal_help_who", labels),
            "comfort": dist(resp, "s1_6_online_form_comfort", labels),
        },
        "et": {
            "contact": dist(resp, "s2_1_et_contact_6m", labels),
            "ctype": dist(resp, "s2_2_contact_type", labels),
            "initiator": dist(resp, "s2_4_contact_initiator", labels),
            "respect": dist(resp, "s2_6_sat_respect", labels),
            "officer_trust": dist(resp, "s3_9_officer_trust", labels),
            "sat_overall": dist(resp, "s3_10_sat_overall", labels),
            "bill_diff": dist(resp, "s9_6_bill_diff_reason", labels),
            "bribe_prev": dist(resp, "s9_9_bribe_prevalence", labels),
            "bribe_asked": dist(resp, "s9_10_bribe_asked", labels),
            "bribe_target": dist(resp, "s9_8_bribe_target_who", labels),
            "visit_target": multi(resp, "s3_11_belief_visit_target", labels),
        },
        "rec": {
            "seen": dist(resp, "s3_1_seen_record", labels),
            "process": multi(resp, "s3_2_assess_process_knowl", labels),
            "land_dir": dist(resp, "s3_4_prior_land_dir", labels),
            "covered_dir": dist(resp, "s3_3_prior_covered_dir", labels),
            "notice": dist(resp, "s3_6_reassess_notice", labels),
            "appeal_aware": multi(resp, "s3_8_appeal_awareness", labels),
            "title": dist(resp, "s3_14_title_status", labels),
            "docs": docs,
            "doc_ease": dist(resp, "s3_17_doc_find_ease", labels),
            "appeal_hist": hist_int(resp, "s9_1_appeal_intention", 0, 10),
            "no_contest": dist(resp, "s9_2_no_contest_reason", labels),
            "reassess_reason": dist(resp, "s3_13_belief_reassess_reason", labels),
        },
        "rs": {
            "mc": mc,
            "post_gap": dist(comp, "s5_1_post_record_gap", labels),
            "reliable": dist(comp, "s5_3_rs_reliable", labels),
            "fair": dist(comp, "s5_4_rs_fair_post", labels),
            "correctable": dist(comp, "s5_5_rs_correctable", labels),
            "err_dir": dist(comp, "s5_6_rs_error_direction", labels),
            "insp_vs_comp": dist(comp, "s5_7_inspector_vs_computer", labels),
            "concerns": multi(comp, "s5_9_tech_concerns", labels),
            "image_conf": dist(comp, "s5_10_image_confidence", labels),
            "accuracy_hist": hist_int(comp, "s5_2_rs_accuracy_oo10", 0, 10),
            "byarm": byarm,
        },
        "wtp": {
            "curve": curve,
            "buy": dist(comp, "s6_1_buy_own_estimate", labels),
            "liability_bins": bins(comp["initial_liability"],
                                   [0, 5000, 10000, 15000, 20000, 30000, 10**9],
                                   lambda a, b: (f"Rs {a//1000}–{b//1000}k" if b < 10**9 else "Rs 30k+")),
            "holdout_hist": hist_int(comp, "s7_holdout_believe", 0, 10),
        },
        "prop": {
            "floors": [{"label": f"{int(k)} floor" + ("s" if k > 1 else ""), "value": int(v)} for k, v in
                       numcol(resp["s13_num_floors"]).dropna().astype(int).value_counts().sort_index().items()],
            "marla_bins": bins(comp["land_area_marla"], [0, 3, 5, 7, 10, 15, 10**9],
                               lambda a, b: (f"{int(a)}–{int(b)} marla" if b < 10**9 else "15+ marla")),
            "covered_bins": bins(comp["covered_area"], [0, 1000, 1500, 2000, 2500, 3000, 10**9],
                                 lambda a, b: (f"{a:,}–{b:,} sq ft" if b < 10**9 else "3,000+ sq ft")),
            "assets": assets,
            "obs_avgs": obs_avgs,
            "dress": dist(resp, "s12_6_obs_western_dress", labels),
            "nervous": dist(resp, "s12_7_obs_nervous", labels),
            "truthful": dist(resp, "s12_8_obs_truthful", labels),
            "distress": dist(comp, "s4t2_5_distress_observed", labels),
        },
        "loc_table": loc_rows,
    }
    return D


def inject(D):
    tpl = TEMPLATE_PATH.read_text(encoding="utf-8")
    payload = json.dumps(D, ensure_ascii=False, separators=(",", ":"))
    out = tpl.replace("/*__DASHBOARD_DATA__*/{}", "const DASHBOARD_DATA = " + payload + ";")
    if out == tpl:
        raise RuntimeError("Data placeholder not found in dashboard_template.html")
    OUT_PATH.write_text(out, encoding="utf-8")


if __name__ == "__main__":
    D = build()
    inject(D)
    m = D["meta"]
    print(f"OK  index.html built — {m['n_complete']} completed / {m['total_target']} target "
          f"({m['pct_complete']}%), {m['n_submissions']} submissions, "
          f"{m['field_days']} field days, updated {m['last_updated']}")
