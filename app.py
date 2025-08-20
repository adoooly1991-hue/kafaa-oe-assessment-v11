
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
profile_key = st.sidebar.selectbox("Industry profile", list(templates.get("profiles",{}).keys()), format_func=lambda k: templates["profiles"][k]["label"] if k in templates.get("profiles",{}) else k)
st.session_state["profile"] = templates.get("profiles",{}).get(profile_key, {})

# ---------- Sidebar ----------
with st.sidebar:
    st.title("🧭 Navigation")
    lang = st.selectbox("Language / اللغة", ["en","ar"], index=0, key="lang")
    nav = st.radio("Go to", ["Welcome","Snapshot","Data Collection","Financial Assessment","Product Selection","VSM Charter","Value Chain","Benchmarks & Rules","Insights & Narratives","Business Case","Export"], index=0, key="nav")

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


elif st.session_state["nav"] == "Data Collection":
    st.subheader("Collect & model your data")
    st.caption("Fill the cards or upload the Kafaa Excel to auto-populate. These fields refine scoring and narratives.")
    up = st.file_uploader("Upload Excel (Kafaa data sheet)", type=["xlsx","xls"])
    if up is not None:
        import pandas as pd, numpy as np
        xls = pd.ExcelFile(up)
        sh = "Material Flow" if "Material Flow" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sh, header=None)
        # Detect columns for 'Process n'
        proc_cols = {}
        for c in range(df.shape[1]):
            for r in range(min(6, df.shape[0])):
                v = str(df.iat[r,c]).strip()
                if v.lower().startswith("process "):
                    proc_cols[v.strip()] = c
        # Helper to find row by key text in col1
        def find_row(label):
            label = label.lower()
            for r in range(df.shape[0]):
                for c in range(min(6, df.shape[1])):
                    v = str(df.iat[r,c]).strip().lower()
                    if label in v:
                        return r
            return None
        row_map = {
            "touchpoints_n": ["n.touch points","n.touchs points","n. touchs points","n. touch points"],
            "ct_min": ["cycle time (min)","cycle time"],
            "process_type": ["process type"],
            "downtime_pct": ["unplanned downtime (%)","unplanned downtime"],
            "defect_pct": ["% defects","defects"],
            "safety_incidents": ["n.safety issues","safety issues"],
            "rework_pct": ["% rework rate","rework"],
            "wip_units_in": ["wip (units)","wip"],
            "push_pull": ["push /pull","push / pull","push /  pull","push /  pull","push / pull","push /  pull","push / pull"],
            "changeover_freq": ["changeover frequency"],
            "changeover_time_min": ["changeover time"],
            "operators_n": ["n.operators","n. operators"]
        }
        rows_found = {}
        for key, keys in row_map.items():
            r = None
            for k in keys:
                r = find_row(k)
                if r is not None: break
            rows_found[key] = r
        # Build steps if needed
        if not st.session_state.get("steps"):
            st.session_state["steps"] = []
        # determine number of processes
        proc_names = sorted(proc_cols.keys(), key=lambda s: int(s.split()[-1]))
        if not proc_names:
            st.warning("Could not find 'Process N' headers. Please review your sheet.")
        else:
            steps = []
            for idx, pname in enumerate(proc_names, start=1):
                col = proc_cols[pname]
                sid = f"P{idx}"
                name = st.session_state.get("steps",[None]*(idx))[:idx][-1].name if st.session_state.get("steps") and len(st.session_state["steps"])>=idx else pname
                # read values
                def _val(row, default=None):
                    if row is None: return default
                    try:
                        v = df.iat[row+1, col]  # value is often on next row
                    except Exception:
                        v = None
                    return v if (v is not None) else default
                new = ProcessStep(
                    id=sid,
                    name=name,
                    ct_sec=float((_val(rows_found["ct_min"]) or 0)*60.0),
                    wip_units_in=float(_val(rows_found["wip_units_in"]) or 0),
                    defect_pct=float(_val(rows_found["defect_pct"]) or 0),
                    rework_pct=float(_val(rows_found["rework_pct"]) or 0),
                    push_pull=str(_val(rows_found["push_pull"]) or "Push").title(),
                    process_type=str(_val(rows_found["process_type"]) or "Manual").title(),
                    distance_m=0.0,
                    layout_moves=int(0),
                    waiting_starved_pct=float(_val(rows_found["downtime_pct"]) or 0),  # initial proxy
                    safety_incidents=int(_val(rows_found["safety_incidents"]) or 0),
                    downtime_pct=float(_val(rows_found["downtime_pct"]) or 0),
                    changeover_freq=float(_val(rows_found["changeover_freq"]) or 0),
                    changeover_time_min=float(_val(rows_found["changeover_time_min"]) or 0),
                    operators_n=float(_val(rows_found["operators_n"]) or 0),
                    touchpoints_n=float(_val(rows_found["touchpoints_n"]) or 0),
                    answers={}
                )
                steps.append(new)
            st.session_state["steps"] = steps
            st.success(f"Imported {len(steps)} processes from Excel.")
    # Manual cards (if no Excel or to refine)
    st.markdown("### Manual entry")
    ensure_default_steps()
    ids = [s.id for s in st.session_state["steps"]]
    tabs = st.tabs(ids)
    for tab, s in zip(tabs, st.session_state["steps"]):
        with tab:
            c1,c2,c3,c4 = st.columns(4)
            with c1:
                s.touchpoints_n = st.number_input(f"{s.id} Touch-points (count)", min_value=0.0, value=float(s.touchpoints_n), step=1.0, key=f"{s.id}-touch")
                s.changeover_freq = st.number_input(f"{s.id} Changeover freq/shift", min_value=0.0, value=float(s.changeover_freq), step=0.5, key=f"{s.id}-cof")
                s.operators_n = st.number_input(f"{s.id} Operators (count)", min_value=0.0, value=float(s.operators_n), step=1.0, key=f"{s.id}-ops")
            with c2:
                s.changeover_time_min = st.number_input(f"{s.id} Changeover time (min)", min_value=0.0, value=float(s.changeover_time_min), step=1.0, key=f"{s.id}-cot")
                s.downtime_pct = st.number_input(f"{s.id} Unplanned downtime (%)", min_value=0.0, value=float(s.downtime_pct), step=1.0, key=f"{s.id}-down")
                s.safety_incidents = st.number_input(f"{s.id} Safety incidents", min_value=0, value=int(s.safety_incidents), step=1, key=f"{s.id}-safe")
            with c3:
                s.rework_pct = st.number_input(f"{s.id} Rework (%)", min_value=0.0, value=float(s.rework_pct), step=0.5, key=f"{s.id}-rew")
                s.defect_pct = st.number_input(f"{s.id} Defects (%)", min_value=0.0, value=float(s.defect_pct), step=0.5, key=f"{s.id}-def2")
                s.wip_units_in = st.number_input(f"{s.id} WIP (units)", min_value=0.0, value=float(s.wip_units_in), step=1.0, key=f"{s.id}-wip2")
            with c4:
                s.ct_sec = st.number_input(f"{s.id} Cycle time (sec)", min_value=0.0, value=float(s.ct_sec), step=5.0, key=f"{s.id}-ct2")
                s.push_pull = st.selectbox(f"{s.id} Mode", ["Push","Pull"], index=0 if s.push_pull=='Push' else 1, key=f"{s.id}-mode2")
                s.process_type = st.selectbox(f"{s.id} Type", ["Manual","Semi-auto","Auto"], index=["Manual","Semi-auto","Auto"].index(s.process_type), key=f"{s.id}-ptype2")
    st.info("Data saved to session. Proceed to ‘Insights & Narratives’ to reflect this data in scoring and observations.")


elif st.session_state["nav"] == "Financial Assessment":
    st.subheader("Financial Assessment — set cost & cash targets")
    up_fin = st.file_uploader("Upload Excel (Financials)", type=["xlsx","xls"], key="up_fin")
    if up_fin is not None:
        import pandas as pd
        x = pd.ExcelFile(up_fin)
        # Try to find a sheet that contains the Financial header row
        target_cols = ["Year","Revenue","COGS","Depreciation","G&A","Financial Expenses","Inventory","Current Assets","Current Liabilities","Sales Target","Budgeted COGS","Budgeted G&A","Budgeted Depreciation","Budgeted Financial Expenses","Targeted Profit"]
        found_df = None
        for sh in x.sheet_names:
            df = pd.read_excel(x, sh)
            headers = [str(c).strip() for c in df.columns]
            if set(target_cols).issubset(set(headers)):
                found_df = df[target_cols].copy()
                break
        if found_df is not None:
            st.session_state["finance_df"] = found_df.head(1)
            st.success("Loaded financials from Excel.")
        else:
            st.warning("Couldn’t find a sheet with the Financial header. Paste the row manually below.")
    st.caption("This section estimates the cost reduction target for the VSM program and a cash/liquidity target from inventory reduction.")
    cols = ["Year","Revenue","COGS","Depreciation","G&A","Financial Expenses","Inventory","Current Assets","Current Liabilities","Sales Target","Budgeted COGS","Budgeted G&A","Budgeted Depreciation","Budgeted Financial Expenses","Targeted Profit"]
    # Pre-fill a single-row editor
    import pandas as pd, numpy as np
    if "finance_df" not in st.session_state:
        st.session_state["finance_df"] = pd.DataFrame([{c: 0.0 for c in cols}])
        st.session_state["finance_df"].loc[0,"Year"] = 2025
    st.write("Enter your latest year figures (absolute amounts in the same currency).")
    fed = st.data_editor(st.session_state["finance_df"], use_container_width=True, num_rows="dynamic")
    st.session_state["finance_df"] = fed

    if st.button("Compute targets", type="primary"):
        r = fed.iloc[0].to_dict()
        # Safe getters
        def g(k): 
            try:
                return float(r.get(k,0) or 0.0)
            except Exception:
                return 0.0
        Revenue = g("Revenue"); COGS = g("COGS"); Dep = g("Depreciation"); GA = g("G&A"); Fin = g("Financial Expenses")
        Inv = g("Inventory"); CA = g("Current Assets"); CL = g("Current Liabilities")
        SalesT = g("Sales Target"); B_COGS = g("Budgeted COGS"); B_GA = g("Budgeted G&A"); B_Dep = g("Budgeted Depreciation"); B_Fin = g("Budgeted Financial Expenses"); TargetProfit = g("Targeted Profit")

        total_costs = COGS + GA + Dep + Fin
        current_profit = Revenue - total_costs
        profit_gap_actual = max(0.0, TargetProfit - current_profit)
        allowable_costs = max(0.0, SalesT - TargetProfit) if SalesT>0 and TargetProfit>0 else None
        required_reduction = max(0.0, total_costs - allowable_costs) if allowable_costs is not None else profit_gap_actual

        # Liquidity metrics (cash proxy): Quick ratio (CA - Inventory)/CL, inventory days, working capital
        quick_ratio = ((CA - Inv) / CL) if CL>0 else None
        inv_days = ((Inv / COGS) * 365.0) if COGS>0 else None
        inv_pct_ca = (Inv / CA) if CA>0 else None
        working_cap = CA - CL
        # Inventory reduction to reach quick ratio 1.0
        inv_reduction_for_qr1 = None
        if CL>0 and CA>0:
            needed_noninv_assets = CL
            current_noninv_assets = max(0.0, CA - Inv)
            gap = max(0.0, needed_noninv_assets - current_noninv_assets)
            inv_reduction_for_qr1 = gap  # cash unlock needed (by reducing inventory or increasing receivables)

        # Suggested allocation: focus on VSM-influenced levers
        # Default shares: COGS 70%, G&A 20%, Financial 10%, Depreciation 0% (fixed)
        shares = {"COGS":0.7, "G&A":0.2, "Financial Expenses":0.1, "Depreciation":0.0}
        # If BUDGETs provided, bias towards areas over budget
        over = {
            "COGS": max(0.0, COGS - B_COGS),
            "G&A": max(0.0, GA - B_GA),
            "Financial Expenses": max(0.0, Fin - B_Fin),
            "Depreciation": max(0.0, Dep - B_Dep)
        }
        if sum(over.values())>0:
            tot_over = sum(over.values())
            # Blend: 60% baseline share + 40% proportional to overrun
            for k in shares:
                shares[k] = 0.6*shares[k] + 0.4*(over[k]/tot_over if tot_over>0 else 0.0)
        allocation = {}
        for k,sh in shares.items():
            amt = required_reduction * sh
            allocation[k] = {"share": sh, "amount": amt, "amount_fmt": f"{amt:,.0f}"}

        notes = []
        notes.append("COGS reduction via waste elimination (defects, waiting, motion, transportation, overprocessing) and yield improvement.")
        notes.append("Inventory actions free cash and may lower financial expenses; target quick ratio ≥ 1.0 as a guardrail.")
        if inv_reduction_for_qr1 is not None and inv_reduction_for_qr1>0:
            notes.append(f"Reduce inventory by ≈ {inv_reduction_for_qr1:,.0f} to reach Quick Ratio 1.0.")
        if inv_days is not None and inv_days>0:
            days_target = max(0.0, inv_days*0.7)  # 30% improvement
            notes.append(f"Reduce Inventory Days from {inv_days:,.0f} to ≈ {days_target:,.0f} (30% improvement).")

        finance = {
            "Year": int(r.get("Year",0) or 0),
            "current_profit": current_profit,
            "Targeted Profit": TargetProfit,
            "profit_gap_actual": profit_gap_actual,
            "required_reduction": required_reduction,
            "quick_ratio": quick_ratio,
            "inventory_days": inv_days,
            "inventory_pct_current_assets": inv_pct_ca,
            "working_capital": working_cap,
            "inv_reduction_for_qr1": inv_reduction_for_qr1,
            "allocation": allocation,
            "notes": notes
        }
        # formatted strings
        finance.update({
            "current_profit_fmt": f"{current_profit:,.0f}",
            "Targeted Profit_fmt": f"{TargetProfit:,.0f}",
            "profit_gap_actual_fmt": f"{profit_gap_actual:,.0f}",
            "required_reduction_fmt": f"{required_reduction:,.0f}",
            "quick_ratio_str": ("{:.2f}".format(quick_ratio) if quick_ratio is not None else "-"),
            "inventory_days_str": ("{:.0f} days".format(inv_days) if inv_days is not None else "-"),
            "inv_pct_ca_str": ("{:.0%}".format(inv_pct_ca) if inv_pct_ca is not None else "-"),
            "inv_reduction_for_qr1_fmt": ("{:,}".format(int(inv_reduction_for_qr1)) if inv_reduction_for_qr1 is not None else "-")
        })
        st.session_state["finance"] = finance

    # Show results if present
    if st.session_state.get("finance"):
        f = st.session_state["finance"]
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Current Profit", f.get("current_profit_fmt","-"))
        c2.metric("Profit Gap (actual)", f.get("profit_gap_actual_fmt","-"))
        c3.metric("Req. Cost Reduction", f.get("required_reduction_fmt","-"))
        c4.metric("Quick Ratio", f.get("quick_ratio_str","-"))
        c5,c6,c7 = st.columns(3)
        c5.metric("Inventory Days", f.get("inventory_days_str","-"))
        c6.metric("Inventory % of CA", f.get("inv_pct_ca_str","-"))
        c7.metric("Inv. reduction for QR=1.0", f.get("inv_reduction_for_qr1_fmt","-"))
        st.markdown("**Suggested allocation:**")
        st.write({k: v["amount_fmt"] for k,v in f.get("allocation",{}).items()})
        if f.get("notes"):
            st.info("\\n".join([f"• {n}" for n in f["notes"]]))


elif st.session_state["nav"] == "Product Selection":
    st.subheader("Product Selection — choose the champion SKU")
    st.caption("Upload your product matrix from Excel or paste the data below. The app scores each product so you can pick the best candidate for the VSM exercise.")
    import pandas as pd, numpy as np, re

    expected_cols = ["Product Name","Cost per Unit (SAR)","Profit per Unit","Total Margin / Unit","Total Cost / Unit","Total Quantity (BOX)","Sales (SAR)","Gross-Margin / Unit  (SAR)","Gross Margin 2020 %","Start Quantity in Inventory Jan 2019","End Quantity in Inventory Dec 2019","Days to Inventory Turnover","Manufacturing Time (Hour)","# of Touching Points - Total"]
    # Uploader
    up = st.file_uploader("Upload Excel (Products)", type=["xlsx","xls"], key="up_products")
    if up is not None:
        x = pd.ExcelFile(up)
        found = None
        for sh in x.sheet_names:
            df = pd.read_excel(x, sh)
            cols = [str(c).strip() for c in df.columns]
            if "Product Name" in cols:
                found = df
                break
        if found is not None:
            st.session_state["products_df_raw"] = found
            st.success(f"Loaded {found.shape[0]} rows from sheet ‘{sh}’.")
        else:
            st.warning("Couldn’t find a sheet with a ‘Product Name’ column. Paste data manually below.")

    # Working copy
    if "products_df" not in st.session_state:
        st.session_state["products_df"] = pd.DataFrame(columns=expected_cols)
    # Attempt to auto-map columns from raw
    if st.session_state.get("products_df_raw") is not None:
        raw = st.session_state["products_df_raw"].copy()
        rename_map = {}
        for c in raw.columns:
            cn = str(c).strip().lower()
            # Basic fuzzy matching by key tokens
            if cn.startswith("product"):
                rename_map[c] = "Product Name"
            elif "cost per unit" in cn and "sar" in cn:
                rename_map[c] = "Cost per Unit (SAR)"
            elif "profit per unit" in cn:
                rename_map[c] = "Profit per Unit"
            elif "total margin" in cn:
                rename_map[c] = "Total Margin / Unit"
            elif ("total cost" in cn) and ("unit" in cn):
                rename_map[c] = "Total Cost / Unit"
            elif ("quantity" in cn) and ("box" in cn):
                rename_map[c] = "Total Quantity (BOX)"
            elif cn.startswith("sales"):
                rename_map[c] = "Sales (SAR)"
            elif ("gross" in cn) and ("unit" in cn):
                rename_map[c] = "Gross-Margin / Unit  (SAR)"
            elif ("gross margin" in cn) and ("%" in cn):
                rename_map[c] = "Gross Margin 2020 %"
            elif ("start" in cn) and ("inventory" in cn):
                rename_map[c] = "Start Quantity in Inventory Jan 2019"
            elif ("end" in cn) and ("inventory" in cn):
                rename_map[c] = "End Quantity in Inventory Dec 2019"
            elif ("inventory" in cn) and ("days" in cn):
                rename_map[c] = "Days to Inventory Turnover"
            elif ("manufacturing time" in cn):
                rename_map[c] = "Manufacturing Time (Hour)"
            elif ("touching points" in cn) or ("touch points" in cn) or ("touchpoints" in cn):
                rename_map[c] = "# of Touching Points - Total"
        dfm = raw.rename(columns=rename_map)
        # Keep only known columns, fill if missing
        for c in expected_cols:
            if c not in dfm.columns:
                dfm[c] = np.nan
        st.session_state["products_df"] = dfm[expected_cols].copy()

    st.markdown("### Edit or paste your product matrix")
    ed = st.data_editor(st.session_state["products_df"], use_container_width=True, num_rows="dynamic")
    st.session_state["products_df"] = ed

    st.markdown("### Scoring")
    st.caption("By default, **highest is best** on all fields. Toggle any field to invert if ‘lower is better’ for your use case (e.g., Cost, Touch-points).")
    num_cols = [c for c in expected_cols if c != "Product Name"]
    cols1, cols2 = st.columns(2)
    invert_default = {"Cost per Unit (SAR)": True, "Total Cost / Unit": True, "Manufacturing Time (Hour)": True, "# of Touching Points - Total": True}
    with cols1:
        invert = {}
        for c in num_cols[:len(num_cols)//2]:
            invert[c] = st.checkbox(f"Invert {c} (lower is better)", value=invert_default.get(c, False), key=f"inv-{c}")
    with cols2:
        for c in num_cols[len(num_cols)//2:]:
            invert[c] = st.checkbox(f"Invert {c} (lower is better)", value=invert_default.get(c, False), key=f"inv-{c}")

    weights = {}
    st.markdown("**Weights (sum auto-normalized)**")
    cw1, cw2 = st.columns(2)
    with cw1:
        for c in num_cols[:len(num_cols)//2]:
            weights[c] = st.number_input(f"Weight: {c}", min_value=0.0, value=1.0, step=0.5, key=f"w-{c}")
    with cw2:
        for c in num_cols[len(num_cols)//2:]:
            weights[c] = st.number_input(f"Weight: {c}", min_value=0.0, value=1.0, step=0.5, key=f"w-{c}")

    if st.button("Rank products", type="primary"):
        df = st.session_state["products_df"].copy()
        # Cast numerics
        for c in num_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        # Build scores: rank percentile for each column (invert if needed), then weighted sum
        total_w = sum(weights.values()) or 1.0
        score_cols = []
        for c in num_cols:
            s = df[c]
            if s.notna().sum() == 0:
                df[f"S_{c}"] = 0.0
            else:
                if invert.get(c, False):
                    r = s.rank(pct=True, ascending=True)  # lower is better
                else:
                    r = s.rank(pct=True, ascending=False) # higher is better
                df[f"S_{c}"] = (r.fillna(0.0)) * (weights[c]/total_w)
            score_cols.append(f"S_{c}")
        df["Total Score"] = df[score_cols].sum(axis=1).round(4)
        df = df.sort_values("Total Score", ascending=False)
        st.session_state["products_ranked"] = df
        # Select champion
        if not df.empty:
            champ_row = df.iloc[0].to_dict()
            champ_notes = []
            top_features = sorted([(c, champ_row.get(c)) for c in num_cols if pd.notna(champ_row.get(c))], key=lambda kv: kv[1] if not invert.get(kv[0], False) else -kv[1], reverse=True)[:3]
            for k,v in top_features:
                champ_notes.append(f"High {k}: {v}")
            st.session_state["champion"] = {"Product Name": champ_row.get("Product Name","-"), "Total Score": float(champ_row.get("Total Score",0.0)), "Notes": "; ".join(champ_notes)}
        st.success("Ranking complete.")

    if st.session_state.get("products_ranked") is not None:
        st.subheader("Results")
        st.dataframe(st.session_state["products_ranked"], use_container_width=True)
        champ = st.session_state.get("champion", {})
        if champ:
            st.info(f"**Champion:** {champ.get('Product Name','-')}  |  Score: {champ.get('Total Score','-')}\n\n{champ.get('Notes','')}")


elif st.session_state["nav"] == "VSM Charter":
    st.subheader("VSM Team Charter — kick-off & sign-off")
    st.caption("Document the official scope, roles, KPIs, financial targets, and team for this VSM exercise. Export as a signed PDF.")

    # Optional: import from sample charter Excel
    up_charter = st.file_uploader("Import from Charter Excel (optional)", type=["xlsx","xls"], key="up_charter")
    charter = st.session_state.get("charter", {})
    if up_charter is not None:
        import pandas as pd, numpy as np, re
        try:
            x = pd.ExcelFile(up_charter)
            sh = "Team Charter" if "Team Charter" in x.sheet_names else x.sheet_names[0]
            df = pd.read_excel(x, sh, header=None)
            def find_value(label):
                label = label.lower()
                for i in range(df.shape[0]):
                    for j in range(df.shape[1]):
                        v = str(df.iat[i,j]).strip().lower() if pd.notna(df.iat[i,j]) else ""
                        if label in v:
                            # prefer right cell or next row same col
                            for jj in range(j+1, df.shape[1]):
                                nv = df.iat[i, jj]
                                if pd.notna(nv) and str(nv).strip()!="":
                                    return str(nv).strip()
                            if i+1<df.shape[0] and pd.notna(df.iat[i+1,j]):
                                return str(df.iat[i+1,j]).strip()
                return None
            charter.update({
                "vs_name": find_value("value stream name") or charter.get("vs_name"),
                "product": find_value("value stream product") or charter.get("product"),
                "start_point": find_value("starting point") or charter.get("start_point"),
                "end_point": find_value("ending point") or charter.get("end_point"),
                "owner": find_value("owner") or charter.get("owner"),
                "exec_sponsor": find_value("executive sponsor") or charter.get("exec_sponsor"),
                "champion_rep": find_value("champion") or charter.get("champion_rep"),
                "facilitator": find_value("facilitator") or charter.get("facilitator"),
                "location": find_value("location") or charter.get("location"),
                "kickoff": find_value("kick-off date") or charter.get("kickoff"),
            })
            st.success("Charter fields pre-filled from Excel. Please review below.")
        except Exception as e:
            st.warning(f"Could not parse charter Excel: {e}")

    # Prefill from other pages (finance + champion)
    finance = st.session_state.get("finance", {})
    champion = st.session_state.get("champion", {})
    if finance:
        charter.setdefault("required_reduction_fmt", finance.get("required_reduction_fmt"))
        charter.setdefault("quick_ratio_str", finance.get("quick_ratio_str"))
        charter.setdefault("inventory_days_str", finance.get("inventory_days_str"))
        charter.setdefault("inv_reduction_for_qr1_fmt", finance.get("inv_reduction_for_qr1_fmt"))
    if champion:
        charter.setdefault("product", champion.get("Product Name"))

    st.markdown("### Header")
    c1,c2,c3 = st.columns(3)
    with c1:
        charter["vs_name"] = st.text_input("Value Stream Name", value=charter.get("vs_name",""))
        charter["product"] = st.text_input("Product", value=charter.get("product",""))
    with c2:
        charter["start_point"] = st.text_input("Starting Point", value=charter.get("start_point",""))
        charter["end_point"] = st.text_input("Ending Point", value=charter.get("end_point",""))
    with c3:
        charter["location"] = st.text_input("Workshop Location", value=charter.get("location",""))
        charter["kickoff"] = st.text_input("Kick-off Date", value=charter.get("kickoff",""))

    st.markdown("### Roles")
    r1,r2,r3,r4 = st.columns(4)
    with r1:
        charter["exec_sponsor"] = st.text_input("Executive Sponsor (CEO)", value=charter.get("exec_sponsor",""))
    with r2:
        charter["owner"] = st.text_input("Value Stream Owner (Client Rep.)", value=charter.get("owner",""))
    with r3:
        charter["champion_rep"] = st.text_input("Value Stream Champion (Service Provider)", value=charter.get("champion_rep",""))
    with r4:
        charter["facilitator"] = st.text_input("Workshop Facilitator", value=charter.get("facilitator",""))

    st.markdown("### Objectives & KPIs")
    default_obj = ""
    if finance:
        default_obj += f"- Achieve cost reduction of {finance.get('required_reduction_fmt','-')} (aligned to targeted profit)\\n"
    if finance and finance.get("inv_reduction_for_qr1_fmt"):
        default_obj += f"- Improve liquidity: Quick Ratio to ≥ 1.0 by reducing inventory ≈ {finance.get('inv_reduction_for_qr1_fmt','-')}\\n"
    default_obj += "- Reduce lead time and WIP via pull/Kanban; improve on-time delivery\\n- Improve quality (defects, rework), safety, and morale"
    charter["objectives"] = st.text_area("Objectives & Success Measures (bullets)", value=charter.get("objectives", default_obj), height=140)

    st.markdown("### Current State Issues & Business Needs")
    charter["issues"] = st.text_area("Issues (bullets)", value=charter.get("issues",""), height=120)

    st.markdown("### Team Members")
    import pandas as pd
    team_df = pd.DataFrame(charter.get("team", [])) if charter.get("team") else pd.DataFrame(columns=["dept","name","contact","role"])
    team_df = st.data_editor(team_df, use_container_width=True, num_rows="dynamic", column_config={
        "dept": "Department",
        "name": "Name",
        "contact": "Contact",
        "role": "Role"
    })
    charter["team"] = team_df.to_dict(orient="records")

    st.markdown("### Approvals")
    a1,a2 = st.columns(2)
    with a1:
        charter["sign_date"] = st.text_input("Sign-off Date", value=charter.get("sign_date",""))
    with a2:
        st.caption("Signatures are captured offline; the PDF contains signature lines.")

    st.session_state["charter"] = charter

    if st.button("Export Charter as PDF", type="primary"):
        from report import export_charter_pdf
        path = export_charter_pdf(charter, "VSM_Charter.pdf", brand_primary=st.session_state.get('brand_primary','#C00000'), logo_path=st.session_state.get('brand_logo_path','assets/kafaa_logo.png'))
        st.success(f"Charter PDF created: {path}")
        with open(path, "rb") as f:
            st.download_button("Download Charter PDF", f, file_name="VSM_Charter.pdf")



elif st.session_state["nav"] == "Value Chain":
    st.subheader("End-to-End Value Chain — Guided Self‑Assessment")
    st.caption("Answer simple questions for each stage. We show best‑practice hints and only ask follow‑ups when your answers suggest a risk.")

    qmap = templates.get("value_chain", {}).get("questions", {})
    conf_levels = templates.get("value_chain", {}).get("confidence", {}).get("levels", [])
    stages = templates.get("value_chain", {}).get("stages", [])
    if not stages:
        st.info("No stages configured in templates.yaml → value_chain.stages")
    else:
        vc_answers = st.session_state.get("vc_answers", {})
        vc_conf = st.session_state.get("vc_confidence", {})
        vc_fu = st.session_state.get("vc_followups", {})
        for stg in stages:
            sid, sname = stg["id"], stg["name"]
            qs = qmap.get(sid, [])
            with st.expander(f"{sname}", expanded=False):
                vc_answers.setdefault(sid, {}); vc_conf.setdefault(sid, {}); vc_fu.setdefault(sid, {})
                for q in qs:
                    cid = f"vc-{sid}-{q['id']}"
                    lbl = q["text"]; helptext = q.get("help","")
                    bm = q.get("benchmark")
                    bm_ref = q.get("bm_ref")
                    if not bm and bm_ref and st.session_state.get("profile"):
                        val = st.session_state["profile"].get("benchmarks",{}).get(bm_ref)
                        if val is not None:
                            unit = "%" if "pct" in bm_ref else (" min" if "min" in bm_ref else "")
                            bm = f"Best practice: {bm_ref.replace('_', ' ').title()} ≈ {val}{unit}"
                    choices = q.get("choices", [])
                    labels = [c["label"] if isinstance(c, dict) else str(c) for c in choices]
                    # Main question
                    sel = st.selectbox(lbl, labels, key=cid, help=helptext + (f"  \n**{bm}**" if bm else ""))
                    # Map to score
                    def _score_from_label(lb):
                        for c in choices:
                            if (isinstance(c, dict) and c.get("label")==lb): return float(c.get("score",0))
                            if (not isinstance(c, dict) and str(c)==lb): return 0.0
                        return 0.0
                    score = _score_from_label(sel)
                    vc_answers[sid][q["id"]] = score
                    # Confidence meter
                    if conf_levels:
                        clabels = [l["label"] for l in conf_levels]
                        csel = st.selectbox("How do you know this?", clabels, index=0, key=cid+"-conf")
                        cf = next((l["factor"] for l in conf_levels if l["label"]==csel), 1.0)
                        vc_conf[sid][q["id"]] = float(cf)
                        st.progress(min(max(cf,0.0),1.0), text=f"Confidence: {cf:.0%}")
                    # Conditional follow-ups
                    fu = q.get("followups", {})
                    if fu and score >= float(fu.get("trigger_score", 3)):
                        st.markdown("**Follow‑up**")
                        vc_fu[sid].setdefault(q["id"], {})
                        for item in fu.get("items", []):
                            iid = f"{cid}-fu-{item['id']}"
                            if item["type"]=="mc":
                                opts = item.get("choices", [])
                                ans = st.selectbox(item["text"], opts, key=iid)
                            elif item["type"]=="text":
                                ans = st.text_input(item["text"], key=iid)
                            elif item["type"]=="num":
                                ans = st.number_input(item["text"], value=0.0, step=1.0, key=iid)
                            else:
                                ans = st.text_input(item["text"], key=iid)
                            vc_fu[sid][q["id"]][item["id"]] = ans
        st.session_state["vc_answers"] = vc_answers
        st.session_state["vc_confidence"] = vc_conf
        st.session_state["vc_followups"] = vc_fu

        if st.button("Compute Value Chain priorities", type="primary"):
            from engine import score_vc_answers
            out = score_vc_answers(vc_answers, templates, vc_confidence=vc_conf, vc_followups=vc_fu)
            vc_summary = []
            for stg in stages:
                sid, sname = stg["id"], stg["name"]
                ranked = out.get(sid,{}).get("ranked", [])
                issues = out.get(sid,{}).get("issues", [])
                conf_i = out.get(sid,{}).get("confidence", 1.0)
                top3 = [(w, sc) for w, sc in ranked[:3] if sc>0]
                vc_summary.append({"stage_name": sname, "top3": top3, "issues": issues, "confidence": conf_i})
            st.session_state["vc_summary"] = vc_summary
            st.success("Value chain priorities updated. Proceed to Insights or export.")
        # Live preview
        if st.session_state.get("vc_summary"):
            st.subheader("Preview: Top wastes per stage")
            for row in st.session_state["vc_summary"]:
                tops = ", ".join([f"{w.title()} ({sc:.1f})" for w,sc in row.get("top3",[])])
                st.write(f"**{row['stage_name']}** → {tops}  \nConfidence: {row.get('confidence',1.0):.0%}")

elif st.session_state["nav"] == "Benchmarks & Rules":
    st.subheader("Benchmarks & Follow‑ups — Admin")
    st.caption("Edit the Value Chain questionnaire, benchmarks, follow‑ups, and confidence levels. Changes apply immediately in this session.")

    import yaml as _yaml
    if "templates_text" not in st.session_state:
        with open("templates.yaml","r",encoding="utf-8") as _f:
            st.session_state["templates_text"] = _f.read()

    st.markdown("**Edit YAML below (advanced users)**")
    text = st.text_area("templates.yaml", value=st.session_state["templates_text"], height=420, label_visibility="collapsed")
    colA, colB = st.columns([1,1])
    if colA.button("Apply changes", type="primary"):
        try:
            parsed = _yaml.safe_load(text)
            st.session_state["templates"] = parsed  # in-memory override
            templates.update(parsed) if isinstance(templates, dict) else None
            st.session_state["templates_text"] = text
            st.success("Templates updated for this session. Re-open Value Chain to see changes.")
        except Exception as e:
            st.error(f"Invalid YAML: {e}")
    if colB.download_button("Download current YAML", data=text, file_name="templates.yaml", mime="text/yaml"):
        pass
    st.info("Tip: To make changes permanent for all users, update templates.yaml in your GitHub repo.")

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


elif st.session_state["nav"] == "Business Case":
    st.subheader("Business Case — quantify potential annual benefit")
    st.caption("Uses your Value Chain results, follow-ups, and the selected industry profile to estimate benefits by waste. Adjust assumptions in templates.yaml → assumptions.")
    from engine import estimate_business_case
    vc_summary = st.session_state.get("vc_summary", [])
    vc_fu = st.session_state.get("vc_followups", {})
    savings = estimate_business_case(vc_summary, templates, vc_followups=vc_fu, assumptions=templates.get("assumptions",{}))
    st.session_state["savings"] = savings
    bw = savings.get("by_waste", {})
    st.write({k: f"{v:,.0f}" for k,v in bw.items()})
    st.metric("Total Estimated Benefit (annual)", f"{savings.get('total',0.0):,.0f}")

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
                finance=st.session_state.get('finance'),
                product_df=st.session_state.get('products_ranked'),
                champion=st.session_state.get('champion'),
                savings=st.session_state.get('savings'),
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
