
import streamlit as st
import pandas as pd
import yaml
import os
from datetime import datetime
from streamlit_lottie import st_lottie

from engine import (
    ProcessStep, score_wastes, make_observation, compute_lead_time,
    build_material_flow_narrative, categorize_theme, get_questionnaire_effects
)
from report import export_observations_pptx, export_observations_pdf

# ---------- App setup ----------
st.set_page_config(page_title="OE Assessment Report Generator", layout="wide")

with open("templates.yaml","r", encoding="utf-8") as f:
    templates = yaml.safe_load(f)

# Branding (locked to Kafaa)
BRAND_PRIMARY = templates.get('brand', {}).get('primary', '#C00000')
BRAND_LOGO = templates.get('brand', {}).get('logo_path', 'assets/kafaa_logo.png')

# Header
st.image(BRAND_LOGO, width=170)
st.title("OE Assessment Report Generator")

# ---------- Sidebar ----------
with st.sidebar:
    st.title("🧭 Navigation")
    lang = st.selectbox("Language / اللغة", ["en","ar"], index=0, key="lang")
    nav = st.radio("Go to", ["Welcome","Snapshot","Value Chain","Insights & Narratives","Export"], index=0, key="nav")

    st.markdown("---")
    st.header("⚙️ Global Settings")
    spacing_mode = st.selectbox("Map spacing uses", ["Effective CT","WIP"], index=0, key="spacing_mode")
    n_steps = st.number_input("Number of process steps", min_value=1, max_value=12, value=st.session_state.get('n_steps',5), step=1, key='n_steps')

    st.header("🏭 Factory / Report Meta")
    factory_name = st.text_input("Factory name", value=st.session_state.get('factory_name','[FactoryName]'), key='factory_name')
    report_year  = st.text_input("Report year", value=st.session_state.get('report_year', str(datetime.now().year)), key='report_year')
    est_cost     = st.text_input("Estimated handling cost (e.g., $250k)", value=st.session_state.get('est_cost','[cost]'), key='est_cost')
    est_sales    = st.text_input("Lost sales opportunity (e.g., $1.2M)", value=st.session_state.get('est_sales','[sales_opportunity]'), key='est_sales')

    st.markdown("---")
    st.header("🎨 Brand Theme")
    st.caption("Kafaa brand is locked for all users (colors and logo).")
    st.write("- Primary: `#C00000`  \n- Secondary: `#FA0000`  \n- Accent: `#FF5B5B`  \n- Text: `#3F3F3F`  \n- Muted: `#7F7F7F`  \n- Background: `#F2F2F2`")
    st.image(BRAND_LOGO, width=140)
    # ensure in session for exporters
    st.session_state['brand_primary'] = BRAND_PRIMARY
    st.session_state['brand_logo_path'] = BRAND_LOGO

    st.markdown("---")
    with st.expander("💾 Save / Load Project"):
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Save snapshot"):
                payload = {
                    "meta": {
                        "factory_name": st.session_state.get("factory_name"),
                        "report_year": st.session_state.get("report_year"),
                        "est_cost": st.session_state.get("est_cost"),
                        "est_sales": st.session_state.get("est_sales"),
                        "spacing_mode": st.session_state.get("spacing_mode"),
                        "lang": st.session_state.get("lang","en"),
                        "n_steps": st.session_state.get("n_steps",5)
                    },
                    "steps": [s.__dict__ for s in st.session_state.get("steps",[])],
                    "vc_summary": st.session_state.get("vc_summary")
                }
                path = "oe_snapshot.json"
                with open(path, "w", encoding="utf-8") as f:
                    import json
                    json.dump(payload, f, ensure_ascii=False, indent=2)
                with open(path, "rb") as f:
                    st.download_button("Download JSON", f, file_name="oe_snapshot.json")
        with c2:
            up = st.file_uploader("Load snapshot (JSON)", type=["json"])
            if up is not None:
                import json
                payload = json.load(up)
                meta = payload.get("meta",{})
                for k,v in meta.items():
                    st.session_state[k] = v
                steps = []
                for sd in payload.get("steps", []):
                    steps.append(ProcessStep(**sd))
                st.session_state["steps"] = steps
                st.session_state["vc_summary"] = payload.get("vc_summary")
                st.success("Snapshot loaded. Use the sidebar to navigate.")

    st.markdown("---")
    with st.expander("🧮 Kanban Sizing"):
        dd = st.number_input("Daily demand (units/day)", min_value=0.0, value=0.0, step=1.0)
        lt = st.number_input("Replenishment lead time (days)", min_value=0.0, value=0.0, step=0.5)
        sf = st.number_input("Safety factor (e.g., 0.2 = 20%)", min_value=0.0, value=0.2, step=0.05)
        cs = st.number_input("Container size (units)", min_value=1.0, value=50.0, step=1.0)
        if st.button("Calculate Kanban cards"):
            import math
            need = dd * lt * (1.0 + sf)
            cards = max(1, math.ceil(need / cs)) if cs>0 else 1
            st.success(f"Recommended Kanban cards: {cards}")

    if st.button("Reset session"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.success("Session cleared.")

# RTL if Arabic
if st.session_state.get('lang','en')=='ar':
    st.markdown('<style>body {direction: rtl; text-align: right;}</style>', unsafe_allow_html=True)

# ---------- State init ----------
def _init_state():
    st.session_state.setdefault("steps", [])
    st.session_state.setdefault("obs_df", pd.DataFrame())
    st.session_state.setdefault("result", None)
    st.session_state.setdefault("vc_summary", None)
    st.session_state.setdefault("material_flow_text", None)
    st.session_state.setdefault("photos", {})
_init_state()

BASIC_COLS = ["id","name","ct_sec","wip_units_in","defect_pct","rework_pct","push_pull","process_type","distance_m","layout_moves","waiting_starved_pct","safety_incidents"]

def steps_to_df(steps):
    rows = []
    for s in steps:
        rows.append({
            "id": s.id, "name": s.name, "ct_sec": s.ct_sec, "wip_units_in": s.wip_units_in,
            "defect_pct": s.defect_pct, "rework_pct": s.rework_pct, "push_pull": s.push_pull,
            "process_type": s.process_type, "distance_m": s.distance_m, "layout_moves": s.layout_moves,
            "waiting_starved_pct": s.waiting_starved_pct, "safety_incidents": s.safety_incidents
        })
    return pd.DataFrame(rows, columns=BASIC_COLS)

def df_to_steps(df, answers_bank=None):
    out = []
    answers_bank = answers_bank or {}
    for _, r in df.iterrows():
        out.append(ProcessStep(
            id=str(r["id"]), name=str(r["name"]), ct_sec=float(r.get("ct_sec",0) or 0.0),
            wip_units_in=float(r.get("wip_units_in",0) or 0.0), defect_pct=float(r.get("defect_pct",0) or 0.0),
            rework_pct=float(r.get("rework_pct",0) or 0.0), push_pull=str(r.get("push_pull","Push") or "Push"),
            process_type=str(r.get("process_type","Manual") or "Manual"), distance_m=float(r.get("distance_m",0) or 0.0),
            layout_moves=int(r.get("layout_moves",0) or 0), waiting_starved_pct=float(r.get("waiting_starved_pct",0) or 0.0),
            safety_incidents=int(r.get("safety_incidents",0) or 0), answers=answers_bank.get(str(r["id"]), {})
        ))
    return out

def ensure_default_steps():
    if not st.session_state["steps"]:
        st.session_state["steps"] = [
            ProcessStep(id=f"P{i}", name=f"Process {i}", ct_sec=60.0, wip_units_in=20.0, defect_pct=1.5,
                        rework_pct=0.0, push_pull="Push", process_type="Manual", distance_m=10.0,
                        layout_moves=1, waiting_starved_pct=5.0, safety_incidents=0, answers={})
            for i in range(1, int(st.session_state.get('n_steps',5))+1)
        ]

# ---------- Pages ----------
if st.session_state["nav"] == "Welcome":
    left, right = st.columns([2,1])
    with left:
        st.subheader("Welcome to your guided Gemba story")
        st.write(
            "This app helps you run an Operational Excellence assessment in minutes—not weeks. "
            "Enter just the basics, answer a few stage questions, and we’ll generate observations, "
            "a current state map, and a Kafaa-branded report—ready for leadership."
        )
        st.markdown(
            "- **Minimal inputs** with confidence markers (● measured / ◐ mixed / ○ inferred)\n"
            "- **Guided questions** per value-chain stage\n"
            "- **Auto observations** + PQCDSM narrative + photos\n"
            "- **Exports** to PPTX/PDF (Kafaa theme)"
        )
    with right:
        try:
            import requests
            url = "https://assets8.lottiefiles.com/packages/lf20_2q9q2kzz.json"
            anim = requests.get(url, timeout=8).json()
            st_lottie(anim, height=220, loop=True, quality="high")
        except Exception:
            st.info("Let's begin → use the left sidebar to move to Snapshot")

elif st.session_state["nav"] == "Snapshot":
    ensure_default_steps()
    st.subheader("Define your process steps")
    mode = st.segmented_control("Choose mode", options=["Simple table","Detailed tabs"], default="Simple table")
    if mode == "Simple table":
        df = steps_to_df(st.session_state["steps"])
        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        st.session_state["steps"] = df_to_steps(edited)
    else:
        ids = [s.id for s in st.session_state["steps"]]
        tabs = st.tabs(ids)
        for tab, s in zip(tabs, st.session_state["steps"]):
            with tab:
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    s.name = st.text_input(f"{s.id} Name", value=s.name, key=f"{s.id}-name")
                    s.ct_sec = st.number_input(f"{s.id} CT (sec)", min_value=0.0, value=float(s.ct_sec), step=5.0, key=f"{s.id}-ct")
                with col2:
                    s.wip_units_in = st.number_input(f"{s.id} WIP", min_value=0.0, value=float(s.wip_units_in), step=1.0, key=f"{s.id}-wip")
                    s.defect_pct = st.number_input(f"{s.id} Defect %", min_value=0.0, value=float(s.defect_pct), step=0.1, key=f"{s.id}-def")
                with col3:
                    s.push_pull = st.selectbox(f"{s.id} Mode", ["Push","Pull"], index=0 if s.push_pull=="Push" else 1, key=f"{s.id}-mode")
                    s.process_type = st.selectbox(f"{s.id} Type", ["Manual","Semi-auto","Auto"], index=["Manual","Semi-auto","Auto"].index(s.process_type), key=f"{s.id}-ptype")
                with col4:
                    s.distance_m = st.number_input(f"{s.id} Distance (m)", min_value=0.0, value=float(s.distance_m), step=1.0, key=f"{s.id}-dist")
                    s.layout_moves = st.number_input(f"{s.id} Layout moves", min_value=0, value=int(s.layout_moves), step=1, key=f"{s.id}-moves")
                s.waiting_starved_pct = st.number_input(f"{s.id} Waiting/starved time (%)", min_value=0.0, value=float(s.waiting_starved_pct), step=0.5, key=f"{s.id}-wait")
                s.safety_incidents = st.number_input(f"{s.id} Safety incidents", min_value=0, value=int(s.safety_incidents), step=1, key=f"{s.id}-safety")

                st.markdown("**Smart questionnaire (optional)** — answers refine severity & narrative")
                qtabs = st.tabs(["Defects","Waiting","Inventory","Transportation","Overprocessing","Motion","Overproduction","Talent","Safety"])
                ans = s.answers or {}
                with qtabs[0]:
                    trend = st.selectbox(f"{s.id} Defect trend (last 4 weeks)", ["(skip)","Rising","Stable","Falling"], index=0, key=f"{s.id}-q-def-trend")
                    if trend != "(skip)":
                        ans.setdefault("defects", {})["trend"] = trend
                with qtabs[1]:
                    freq = st.selectbox(f"{s.id} Starved/blocked frequency", ["(skip)","Frequent","Occasional","Rare"], index=0, key=f"{s.id}-q-wait-freq")
                    if freq != "(skip)":
                        ans.setdefault("waiting", {})["frequency"] = freq

                st.markdown("**Photos (optional)** — attach evidence per waste; will appear on PPTX detail slides")
                photo_tabs = st.tabs(["Defects","Waiting","Inventory","Transportation","Overprocessing","Motion","Overproduction","Talent","Safety"])
                waste_keys = ["defects","waiting","inventory","transportation","overprocessing","motion","overproduction","talent","safety"]
                for t, wkey in zip(photo_tabs, waste_keys):
                    with t:
                        files = st.file_uploader(f"{s.id} — {wkey.title()} photos", type=["png","jpg","jpeg","webp"], accept_multiple_files=True, key=f"{s.id}-up-{wkey}")
                        if files:
                            save_dir = os.path.join("uploads", s.id, wkey)
                            os.makedirs(save_dir, exist_ok=True)
                            saved_paths = []
                            for uf in files:
                                fn = uf.name.replace(" ", "_")
                                path = os.path.join(save_dir, fn)
                                with open(path, "wb") as out:
                                    out.write(uf.getbuffer())
                                saved_paths.append(path)
                            st.session_state["photos"].setdefault((s.id, wkey), [])
                            existing = st.session_state["photos"][(s.id, wkey)]
                            for p in saved_paths:
                                if p not in existing:
                                    existing.append(p)
                            st.success(f"Saved {len(saved_paths)} file(s).")
                        for p in st.session_state.get("photos", {}).get((s.id, wkey), [])[:3]:
                            st.image(p, caption=os.path.basename(p), use_column_width=True)
                s.answers = ans

elif st.session_state["nav"] == "Value Chain":
    st.subheader("End-to-End Stages")
    vc = templates.get("value_chain", {})
    vc_stages = vc.get("stages", [])
    if not vc_stages:
        st.info("No stages configured in templates.yaml → value_chain.stages")
    else:
        tabs = st.tabs([s["name"] for s in vc_stages])
        vc_scores = {}
        for tab, stage in zip(tabs, vc_stages):
            with tab:
                sid = stage["id"]
                vc_scores[sid] = vc_scores.get(sid, {})
                q1 = st.selectbox(f"{stage['name']}: Supply/Flow stability", ["(skip)","Poor","Average","Good"], index=0, key=f"vc-{sid}-q1")
                if q1 == "Poor":
                    vc_scores[sid]["waiting"] = vc_scores[sid].get("waiting",0)+1.0
                if q1 == "Average":
                    vc_scores[sid]["waiting"] = vc_scores[sid].get("waiting",0)+0.5
                q2 = st.selectbox(f"{stage['name']}: Quality leakage", ["(skip)","High","Medium","Low"], index=0, key=f"vc-{sid}-q2")
                if q2 == "High":
                    vc_scores[sid]["defects"] = vc_scores[sid].get("defects",0)+1.0
                if q2 == "Medium":
                    vc_scores[sid]["defects"] = vc_scores[sid].get("defects",0)+0.5
                st.caption("Top-3 (live)")
                ranked = sorted(vc_scores[sid].items(), key=lambda kv: kv[1], reverse=True)
                st.write(ranked[:3])

        vc_summary = []
        common = templates.get("value_chain", {}).get("common_issues", {})
        for stg in vc_stages:
            sid = stg["id"]
            ranked = sorted(vc_scores.get(sid, {}).items(), key=lambda kv: kv[1], reverse=True)
            top3 = [(m, sc) for m, sc in ranked[:3] if sc > 0]
            issues = []
            if top3:
                issues.extend(common.get(sid, {}).get(top3[0][0], [])[:2])
            vc_summary.append({"stage_name": stg["name"], "top3": top3, "issues": issues})
        st.session_state["vc_summary"] = vc_summary

elif st.session_state["nav"] == "Insights & Narratives":
    st.subheader("Generate insights")
    if st.button("Run analysis", type="primary"):
        steps = st.session_state["steps"] or []
        if not steps:
            st.warning("Please add steps in Snapshot first.")
            st.stop()
        result = compute_lead_time(steps, available_time_sec=8*3600.0)
        st.session_state["result"] = result

        rows = []
        for s in steps:
            wres = score_wastes(s, templates["thresholds"], templates=templates)
            for waste in ["defects","waiting","inventory","overproduction","transportation","motion","overprocessing","talent","safety"]:
                row = make_observation(s, waste, wres, templates, templates["thresholds"])
                if row:
                    rows.append(row)
        obs = pd.DataFrame(rows)
        if not obs.empty:
            obs = obs.sort_values(["rpn_pct","score_0_5"], ascending=False).reset_index(drop=True)
            id_to_step = {s.id:s for s in steps}
            ev_list=[]; mk_list=[]; tip_list=[]
            for r in obs.itertuples(index=False):
                stp = id_to_step.get(getattr(r,'step_id', None), None) or next((s for s in steps if s.name==r.step_name), None)
                w = r.waste
                primary = False
                if stp:
                    if w=='defects': primary = (stp.defect_pct or 0)>0
                    elif w=='waiting': primary = (stp.waiting_starved_pct or 0)>0
                    elif w=='inventory': primary = (stp.wip_units_in or 0)>0
                    elif w=='transportation': primary = (stp.distance_m or 0)>0 or (stp.layout_moves or 0)>0
                    elif w=='motion': primary = True if (stp.process_type or 'Manual') else False
                    elif w=='overprocessing': primary = (stp.rework_pct or 0)>0
                    elif w=='overproduction': primary = True
                    elif w=='safety': primary = (stp.safety_incidents or 0)>0
                dlt,_ = get_questionnaire_effects(stp, templates, w) if stp else (0.0,[])
                if primary and dlt>0:
                    ev='Mixed'
                elif primary:
                    ev='Measured'
                else:
                    ev='Inferred'
                mk = '●' if ev=='Measured' else ('◐' if ev=='Mixed' else '○')
                tip = 'Measured: direct metrics' if ev=='Measured' else ('Mixed: metrics + questionnaire' if ev=='Mixed' else 'Inferred: questionnaire/heuristics')
                ev_list.append(ev); mk_list.append(mk); tip_list.append(tip)
            obs['evidence'] = ev_list; obs['evidence_marker'] = mk_list; obs['evidence_note'] = tip_list
        st.session_state["obs_df"] = obs

        mf = build_material_flow_narrative(steps, templates, factory_name, report_year, est_cost, est_sales)
        st.session_state["material_flow_text"] = mf
        st.success("Insights generated.")

    if st.session_state.get("result"):
        res = st.session_state["result"]
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Total Lead Time (sec)", f"{int(res['lead_time_sec'])}")
        with c2:
            st.metric("Bottleneck CT (sec)", f"{int(res['ct_bottleneck_sec'])}")
        with c3:
            st.metric("Observations", f"{len(st.session_state.get('obs_df', pd.DataFrame()))}")

    obs = st.session_state.get("obs_df", pd.DataFrame())
    if not obs.empty:
        st.dataframe(obs, use_container_width=True)
        obs["theme_code"] = obs["waste"].apply(lambda w: categorize_theme(w)[0])
        theme_order = [("P","Production"),("Q","Quality"),("C","Cost"),("D","Delivery"),("S","Safety"),("M","Morale")]
        st.subheader("Narrative by Theme (PQCDSM)")
        for code, name in theme_order:
            grp = obs[obs["theme_code"]==code]
            if grp.empty:
                continue
            with st.expander(f"{code} — {name}", expanded=False):
                for idx, r in enumerate(grp.itertuples(index=False), start=1):
                    num = f"{code}-{idx}"
                    st.markdown(f"**{num}: {r.step_name} — {r.waste.title()}**  \\n{getattr(r,'observation','')}")

        st.subheader("Material Flow Narrative")
        st.write(st.session_state.get("material_flow_text",""))

elif st.session_state["nav"] == "Export":
    colA, colB = st.columns(2)
    with colA:
        st.caption('Using Kafaa PPTX master by default. (assets/kafaa_guideline.pptx)')
        template_master = None
        if st.button("Export PPTX", type="primary"):
            obs_df = st.session_state.get("obs_df", pd.DataFrame())
            if obs_df.empty:
                st.warning("Generate insights first.")
                st.stop()
            steps = st.session_state.get("steps", [])
            perstep_top2 = {}
            for s in steps:
                w2 = score_wastes(s, templates["thresholds"], templates=templates)
                ranked = sorted(list(w2["scores"].items()), key=lambda kv: kv[1], reverse=True)
                perstep_top2[s.id] = [(name,score) for name,score in ranked if score>0][:2]
            result = st.session_state.get("result", {"by_step":{}})
            ct_eff_map = {sid: result.get("by_step",{}).get(sid,{}).get("ct_eff_sec",0.0) for sid in result.get("by_step",{}).keys()}
            template_path = None  # use default from templates.yaml
            path = export_observations_pptx(
                obs_df, "oe_assessment.pptx",
                steps=steps, perstep_top2=perstep_top2,
                spacing_mode=st.session_state.get("spacing_mode","Effective CT"),
                ct_eff_map=ct_eff_map,
                vc_summary=st.session_state.get("vc_summary"),
                material_flow_text=st.session_state.get("material_flow_text"),
                photos=st.session_state.get("photos"),
                template_path=template_path,
                lang=st.session_state.get("lang","en"),
                i18n=templates.get("i18n",{}),
                brand_primary=st.session_state.get('brand_primary',BRAND_PRIMARY),
                logo_path=st.session_state.get('brand_logo_path',BRAND_LOGO)
            )
            st.success(f"PPTX created: {path}")
            with open(path, "rb") as f:
                st.download_button("Download PPTX", f, file_name="OE_Assessment_Report.pptx")
    with colB:
        if st.button("Export PDF"):
            obs_df = st.session_state.get("obs_df", pd.DataFrame())
            if obs_df.empty:
                st.warning("Generate insights first.")
                st.stop()
            path = export_observations_pdf(
                obs_df, "oe_assessment.pdf",
                brand_primary=st.session_state.get('brand_primary',BRAND_PRIMARY),
                logo_path=st.session_state.get('brand_logo_path',BRAND_LOGO)
            )
            st.success(f"PDF created: {path}")
            with open(path, "rb") as f:
                st.download_button("Download PDF", f, file_name="OE_Assessment_Report.pdf")
