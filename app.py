"""
PFFF v13 — Final Streamlit App
================================
All bugs fixed. Production-ready. Matches Colab output exactly.

KEY FIXES:
1. Switching Values: Dual-anchor (DPR vs P50) — shows Phantom Safety bias
2. Zero-stress toggle: shows calibration proof inline, no stale charts
3. Plotly colors: ALL use rgba() format — no hex alpha (#FF000011 crash fixed)
4. Delay SV correctly shows "Already failed" for RED projects
5. LA% sensitivity chart explains the IRC SP:30 logic
6. All 7 projects in batch with correct FI ordering
"""

import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from scipy.optimize import brentq
import io, json, warnings
warnings.filterwarnings("ignore")

try:
    from pfff_engine import (
        PROJECTS, MODES, HURDLES,
        compute_scn, run_mcs, simulate_mode,
        spearman_tornado, rcf_acid_test, eirr_iter,
        fi_color, verdict, compute_dual_sv,
    )
except ImportError as e:
    st.error(f"pfff_engine.py not found in the same folder.\nError: {e}")
    st.stop()

st.set_page_config(
    page_title="PFFF v13 — NHAI DPR Fragility Auditor",
    page_icon="🏛️", layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif;}
.block-container{padding-top:0.4rem;padding-bottom:0.4rem;}
.fi-badge{border-radius:10px;padding:16px 20px;text-align:center;border:2px solid;}
.fi-big{font-size:3rem;font-weight:700;line-height:1.1;}
.fi-sub{font-size:0.95rem;margin-top:4px;}
.kpi-box{background:#f8f9fa;border-radius:8px;padding:12px 8px;
         border-left:4px solid #dee2e6;text-align:center;margin-bottom:6px;}
.kpi-val{font-size:1.65rem;font-weight:700;line-height:1.2;}
.kpi-lbl{font-size:0.73rem;color:#6c757d;}
.sv-dpr{background:#FFF3CD;border-radius:6px;padding:10px 14px;margin:4px 0;border-left:4px solid #856404;}
.sv-p50{background:#F8D7DA;border-radius:6px;padding:10px 14px;margin:4px 0;border-left:4px solid #842029;}
.sv-ok {background:#D1E7DD;border-radius:6px;padding:10px 14px;margin:4px 0;border-left:4px solid #198754;}
.note  {background:#e8f4fd;border-left:4px solid #0d6efd;border-radius:6px;
        padding:10px 14px;font-size:0.86rem;color:#0c3c60;margin:6px 0;}
.bias-box{background:#F8D7DA;border-left:5px solid #842029;border-radius:8px;
          padding:14px;margin:8px 0;font-size:0.9rem;}
.zs-box{background:#D1E7DD;border-left:5px solid #198754;border-radius:8px;
        padding:12px;margin:8px 0;font-size:0.88rem;}
</style>
""", unsafe_allow_html=True)


def _fc(fi):
    return "#198754" if fi<25 else "#856404" if fi<50 else "#842029"

def _bg(fi):
    return "#D1E7DD" if fi<25 else "#FFF3CD" if fi<50 else "#F8D7DA"

def _vt(fi):
    return "GREEN" if fi<25 else "AMBER" if fi<50 else "RED"

# rgba colors for Plotly (no #RRGGBBAA format — crashes Streamlit Cloud)
RGBA = {
    "red_fill":  "rgba(220,53,69,0.08)",
    "red_zone":  "rgba(220,53,69,0.12)",
    "amber_fill":"rgba(255,193,7,0.07)",
    "green_fill":"rgba(25,135,84,0.07)",
    "p50_line":  "rgba(13,110,253,1)",
    "dpr_line":  "rgba(33,37,41,1)",
}


@st.cache_data(show_spinner=False, ttl=None)
def _sim(pj, mode, n):
    p=json.loads(pj); scn=compute_scn(p); samp=run_mcs(p,scn,n)
    res=simulate_mode(p,scn,samp,mode,n)
    torn=spearman_tornado(p,scn,samp,res["eirr_arr"])
    rcf=rcf_acid_test(p,scn,samp,res["fi_p"])
    ep=res["eirr_arr"]*100; p50=np.percentile(ep,50)
    svs=compute_dual_sv(p,scn,p50)
    return res, scn, samp, torn, rcf, svs, p50


@st.cache_data(show_spinner=False, ttl=None)
def _zs_proof(pj):
    p=json.loads(pj); scn=compute_scn(p)
    zs=eirr_iter(p,scn,p["civil_cr"],0.0,p["yr1_aadt"],p["growth"],1.0,1.0)*100
    return zs, abs(zs-p["dpr_eirr"])<0.05


# ══════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🏛️ PFFF v13")
    st.caption("Probabilistic Feasibility Fragility Framework\nM.BEM Thesis | SPA Delhi 2024 | Varshni M S")
    st.divider()

    proj_key = st.selectbox("Select Project Template",
                             list(PROJECTS.keys()),
                             format_func=lambda k: f"{k} — {PROJECTS[k]['name'][:32]}")
    if st.button("📂 Load Project", use_container_width=True, type="primary"):
        st.session_state["p"] = dict(PROJECTS[proj_key])
    if "p" not in st.session_state:
        st.session_state["p"] = dict(PROJECTS["P2"])
    p = st.session_state["p"]
    st.divider()

    n_iter   = st.select_slider("Monte Carlo Iterations",[1000,2000,5000,10000],value=5000)
    sim_mode = st.selectbox("Procurement Mode",MODES,index=MODES.index(p.get("dpr_mode","EPC")))
    st.divider()

    with st.expander("📈 Economic Parameters", expanded=True):
        p["dpr_eirr"]  = st.number_input("DPR EIRR (%)",value=float(p["dpr_eirr"]),step=0.1)
        p["cost_sens"] = st.number_input("Cost Sensitivity (pp/1%)",value=float(p.get("cost_sens",0.15)),step=0.01)
        p["traf_sens"] = st.number_input("Traffic Sensitivity (pp/1%)",value=float(p.get("traf_sens",0.20)),step=0.01)
        has_firr=st.checkbox("Has FIRR",value=(p.get("dpr_firr") not in (None,0)))
        p["dpr_firr"]=st.number_input("FIRR (%)",value=float(p.get("dpr_firr") or 12.0),step=0.1) if has_firr else None
        has_eq=st.checkbox("Has Equity IRR",value=(p.get("dpr_eq") not in (None,0)))
        p["dpr_eq"]=st.number_input("Equity IRR (%)",value=float(p.get("dpr_eq") or 15.0),step=0.1) if has_eq else None

    with st.expander("💰 Costs & Traffic", expanded=False):
        p["civil_cr"]=st.number_input("Civil Cost (₹ Cr)",value=float(p["civil_cr"]),step=10.0)
        p["la_cr"]   =st.number_input("LA Cost (₹ Cr)",value=float(p["la_cr"]),step=10.0)
        p["om_cr"]   =st.number_input("O&M Yr1 (₹ Cr)",value=float(p.get("om_cr",20.0)))
        p["scale_cr"]=p["civil_cr"]
        p["base_aadt"]=st.number_input("Base AADT",value=int(p["base_aadt"]))
        p["yr1_aadt"] =st.number_input("Year-1 AADT (DPR)",value=int(p["yr1_aadt"]))
        p["growth"]   =st.number_input("Growth Rate",value=float(p.get("growth",0.065)),step=0.005)
        p["dpr_yr"]   =st.number_input("DPR Year",value=int(p.get("dpr_yr",2020)),step=1,min_value=1990,max_value=2030)
        p["survey_yr"]=st.number_input("Survey Year",value=int(p.get("survey_yr",2019)),step=1,min_value=1990,max_value=2030)
        p["survey_indep"]=st.checkbox("Independent Survey",value=bool(p.get("survey_indep",False)))

    with st.expander("🏗️ Risk Conditioners (SCN)", expanded=True):
        p["la_pct"]=st.slider("LA% Complete at DPR",0,100,int(p.get("la_pct",50)),
            help="Higher → lower delay risk AND lower LA cost overrun. EIRR unaffected (IRC SP:30: LA is transfer payment).")
        p["geotech"]=st.select_slider("Geotech Quality",["DESKTOP","PARTIAL","COMPLETE"],value=p.get("geotech","PARTIAL"))
        p["contractor"]=st.select_slider("Contractor",["STRESSED","ADEQUATE","STRONG"],value=p.get("contractor","ADEQUATE"))
        p["community"]=st.select_slider("Community Risk",["LOW","LOW_MEDIUM","MEDIUM","HIGH","EXTREME"],value=p.get("community","MEDIUM"))
        p["terrain"]=st.selectbox("Terrain",["PLAIN","ROLLING","COASTAL_ROLLING","HILLY","MIXED_MOUNTAIN","MOUNTAIN"],
                                   index=["PLAIN","ROLLING","COASTAL_ROLLING","HILLY","MIXED_MOUNTAIN","MOUNTAIN"].index(p.get("terrain","PLAIN")))
        p["forest_clr"]=st.selectbox("Forest Clearance",
                                      ["NONE","CLEARED","EIA_PENDING","NOT_APPLIED","PENDING","STAGE_II","BLOCKED"],
                                      index=["NONE","CLEARED","EIA_PENDING","NOT_APPLIED","PENDING","STAGE_II","BLOCKED"].index(p.get("forest_clr","NONE")))
        p["crossings"]=st.selectbox("Major Crossings",["LOW","MODERATE","HIGH","VERY_HIGH"],
                                     index=["LOW","MODERATE","HIGH","VERY_HIGH"].index(p.get("crossings","LOW")))
        p["network"]=st.selectbox("Network Type",["STANDALONE","FEEDER","CORRIDOR_LINK"],
                                   index=["STANDALONE","FEEDER","CORRIDOR_LINK"].index(p.get("network","FEEDER")))
        p["proj_type"]=st.selectbox("Project Type",["GREENFIELD","BROWNFIELD"],
                                     index=["GREENFIELD","BROWNFIELD"].index(p.get("proj_type","GREENFIELD")))
        p["forest_pct"]=st.number_input("Forest Area (%)",value=float(p.get("forest_pct",0.0)))

    st.session_state["p"]=p

# ══════════════════════════════════════════════════════════════════════
# RUN
# ══════════════════════════════════════════════════════════════════════
pj = json.dumps(p, default=str)
with st.spinner(f"Running {n_iter:,} iterations…"):
    res, scn, samp, tornado, rcf, svs, p50 = _sim(pj, sim_mode, n_iter)

ep=res["eirr_arr"]*100
fi=res["fi_p"]
p10,p20,p80,p90=np.percentile(ep,10),np.percentile(ep,20),np.percentile(ep,80),np.percentile(ep,90)
bias_gap = p["dpr_eirr"] - p50
already_failed = p50 < 12.0

# ══════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════
col_t, col_b = st.columns([4,1])
with col_t:
    st.markdown(f"## 🏛️ {p['name']}")
    st.caption(f"Mode: **{sim_mode}** (DPR: **{p['dpr_mode']}**) · Survey age: **{scn['survey_age']}yr** · "
               f"DPR EIRR: **{p['dpr_eirr']:.2f}%** · P50 Simulated: **{p50:.2f}%** · "
               f"Optimism Bias: **{bias_gap:+.2f}pp** · {n_iter:,} iterations")
with col_b:
    st.markdown(f"""<div class='fi-badge' style='background:{_bg(fi)};border-color:{_fc(fi)}'>
    <div class='fi-big' style='color:{_fc(fi)}'>{fi:.1f}%</div>
    <div class='fi-sub' style='color:{_fc(fi)}'>{_vt(fi)} — {["Approve","Conditional","Return DPR"][["GREEN","AMBER","RED"].index(_vt(fi))]}</div>
    </div>""", unsafe_allow_html=True)

# ── Bias alert ─────────────────────────────────────────────────────────
if abs(bias_gap) > 2:
    st.markdown(f"""<div class='bias-box'>
    <b>⚠ Optimism Bias Detected: {bias_gap:+.2f}pp</b><br>
    The consultant's DPR states EIRR = <b>{p['dpr_eirr']:.2f}%</b>. 
    PFFF's probabilistic simulation gives a P50 (median) of <b>{p50:.2f}%</b>.
    {"The project is <b>already below the 12% viability threshold at median outcomes</b>. The DPR's switching values are phantom safety — they assume the consultant's optimistic anchor." if already_failed else
    f"The project is {p50-12:.2f}pp above hurdle at P50, but with {bias_gap:.1f}pp less margin than the DPR claims."}
    </div>""", unsafe_allow_html=True)

# ── Zero-stress toggle ──────────────────────────────────────────────────
show_zs = st.toggle("🟢 Show Zero-Stress Calibration Proof")
if show_zs:
    zs_val, zs_ok = _zs_proof(pj)
    st.markdown(f"""<div class='zs-box'>
    <b>✅ Zero-Stress Calibration Proof</b><br>
    When the model is fed <i>exactly</i> the DPR's stated values (no overrun, no delay, traffic = DPR forecast, V10=V11=1.0):<br>
    → Simulated EIRR = <b>{zs_val:.4f}%</b> &nbsp;|&nbsp; DPR Stated EIRR = <b>{p['dpr_eirr']:.4f}%</b> &nbsp;|&nbsp; 
    {'✅ <b>PASS</b> (Δ < 0.05pp)' if zs_ok else f'⚠️ <b>DEVIATION</b> (Δ={abs(zs_val-p["dpr_eirr"]):.3f}pp)'}<br>
    <small>This confirms the model is not biased against projects. The FI ({fi:.1f}%) reflects what happens when 
    <i>realistic uncertainty</i> is applied to the DPR's assumptions — not a pre-determined rejection.</small>
    </div>""", unsafe_allow_html=True)
    st.markdown("")

# ── KPI row ────────────────────────────────────────────────────────────
k1,k2,k3,k4,k5,k6 = st.columns(6)
def _kpi(col, val, lbl, color, sub=""):
    col.markdown(f"""<div class='kpi-box' style='border-left-color:{color}'>
    <div class='kpi-val' style='color:{color}'>{val}</div>
    <div class='kpi-lbl'>{lbl}</div>
    <div class='kpi-lbl'>{sub}</div>
    </div>""", unsafe_allow_html=True)

_kpi(k1, f"{fi:.1f}%", "FI Primary", _fc(fi), _vt(fi))
_kpi(k2, f"{res['fi_eirr']:.1f}%", "FI EIRR", _fc(res['fi_eirr']), "Hurdle 12%")
fi_f = f"{res['fi_firr']:.1f}%" if not np.isnan(res['fi_firr']) else "N/A"
_kpi(k3, fi_f, "FI FIRR", _fc(res['fi_firr']) if not np.isnan(res['fi_firr']) else "#6c757d", "Hurdle 10%")
fi_e = f"{res['fi_eq']:.1f}%" if not np.isnan(res['fi_eq']) else "N/A"
eq_h = res.get('hurdle_eq')
_kpi(k4, fi_e, "FI Equity", _fc(res['fi_eq']) if not np.isnan(res['fi_eq']) else "#6c757d", f"Hurdle {eq_h*100:.0f}%" if eq_h else "N/A")
_kpi(k5, f"{p50:.2f}%", "P50 EIRR", "#198754" if p50>=12 else "#842029", f"DPR: {p['dpr_eirr']:.2f}%")
_kpi(k6, f"{bias_gap:+.2f}pp", "Optimism Bias", "#842029" if abs(bias_gap)>1 else "#198754", "DPR − P50 simulated")

st.divider()

# ══════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════
tab1,tab2,tab3,tab4,tab5 = st.tabs([
    "📊 IRR Distributions",
    "🎯 Fragility Drivers",
    "🔑 Switching Values",
    "📋 All 7 Projects",
    "💾 Export",
])

# ─────────────────────────────────────────────────────────────────────
# TAB 1 — IRR DISTRIBUTIONS
# ─────────────────────────────────────────────────────────────────────
with tab1:
    def _hist(arr, hurdle, color_hex, title, dpr_val=None):
        valid = arr[~np.isnan(arr)]*100 if arr is not None else np.array([])
        if len(valid)<10:
            f=go.Figure(); f.add_annotation(text=f"{title}<br>Not applicable for {sim_mode} mode",
                xref="paper",yref="paper",x=0.5,y=0.5,showarrow=False,font=dict(size=13,color="#6c757d"))
            f.update_layout(height=400,plot_bgcolor="#FAFAFA",paper_bgcolor="white",xaxis_visible=False,yaxis_visible=False)
            return f
        fi_v=np.sum(valid<hurdle*100)/len(valid)*100
        p20_,p50_,p80_=np.percentile(valid,20),np.percentile(valid,50),np.percentile(valid,80)
        f=go.Figure()
        f.add_vrect(x0=min(valid)-3,x1=hurdle*100,fillcolor=RGBA["red_fill"],line_width=0,
                    annotation_text="Below hurdle",annotation_position="top left",
                    annotation_font_color="#842029",annotation_font_size=9)
        f.add_trace(go.Histogram(x=valid,nbinsx=55,name="Simulated",
                                  marker_color=color_hex,marker_line=dict(color="white",width=0.4),opacity=0.85))
        f.add_vline(x=hurdle*100,line_dash="dash",line_color="#DC3545",line_width=2.5,
                    annotation_text=f"Hurdle {hurdle*100:.0f}%",annotation_font_color="#DC3545")
        if dpr_val:
            f.add_vline(x=dpr_val,line_dash="solid",line_color="#212529",line_width=2,
                        annotation_text=f"DPR {dpr_val:.1f}%",annotation_position="top right")
        f.add_vline(x=p50_,line_dash="dot",line_color="#0D6EFD",line_width=2,
                    annotation_text=f"P50 {p50_:.1f}%",annotation_position="top left")
        for pv,pn,pc in [(p20_,"P20","#FFC107"),(p80_,"P80","#198754")]:
            f.add_vline(x=pv,line_dash="longdash",line_color=pc,line_width=1.2)
        if dpr_val and dpr_val!=p50_:
            bias=dpr_val-p50_
            f.add_annotation(x=(dpr_val+p50_)/2,y=0,xref="x",yref="paper",
                              text=f"Bias\n{bias:+.1f}pp",showarrow=False,
                              font=dict(size=10,color="#842029"),bgcolor="white",
                              bordercolor="#842029",borderwidth=1.5,borderpad=4)
        f.add_annotation(text=f"<b>FI = {fi_v:.1f}%</b><br>{_vt(fi_v)}",
                         xref="paper",yref="paper",x=0.02,y=0.96,showarrow=False,
                         bgcolor=_bg(fi_v),bordercolor=_fc(fi_v),borderwidth=1.5,borderpad=5,
                         font=dict(size=11,color=_fc(fi_v)))
        f.update_layout(title=f"<b>{title}</b>",height=410,plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                        bargap=0.04,showlegend=False,
                        xaxis=dict(title="IRR (%)",gridcolor="#EEEEEE"),
                        yaxis=dict(title="Frequency",gridcolor="#EEEEEE"),
                        margin=dict(l=50,r=50,t=50,b=40))
        return f

    c1,c2,c3=st.columns(3)
    with c1: st.plotly_chart(_hist(res["eirr_arr"],HURDLES["EIRR"],"#17A589","EIRR — Society's View (12% hurdle)",p["dpr_eirr"]),use_container_width=True)
    with c2: st.plotly_chart(_hist(res["firr_arr"] if not np.all(np.isnan(res["firr_arr"])) else None,
                                    HURDLES["FIRR"],"#8E44AD","FIRR — Lender's View (10% hurdle)",p.get("dpr_firr")),use_container_width=True)
    with c3:
        eq_h_v=res.get("hurdle_eq") or HURDLES["EQ_BOT"]
        st.plotly_chart(_hist(res["eq_arr"] if not np.all(np.isnan(res["eq_arr"])) else None,
                               eq_h_v,"#2471A3",f"Equity IRR — Concessionaire ({eq_h_v*100:.0f}% hurdle)",p.get("dpr_eq")),use_container_width=True)

    st.markdown("""<div class='note'>
    <b>Reading these charts:</b> Each bar = frequency of that IRR in 5,000 simulations.
    The <b>black solid line</b> = DPR's stated value (consultant's inside view).
    The <b>blue dotted line</b> = P50 simulated (PFFF's outside view).
    The gap between them = <b>Optimism Bias</b>.
    Red shaded zone = failure region.
    </div>""", unsafe_allow_html=True)

    # Percentile table
    st.markdown("#### EIRR Percentile Summary")
    df_pct=pd.DataFrame({
        "Percentile":["P10","P20","P50 (central)","P80","P90","DPR Stated"],
        "EIRR (%)":  [round(x,2) for x in [p10,p20,p50,p80,p90,p["dpr_eirr"]]],
        "vs Hurdle 12%":[f"{x-12:+.2f}pp" for x in [p10,p20,p50,p80,p90,p["dpr_eirr"]]],
        "Meaning":["10% chance EIRR is below this","20% chance EIRR is below this",
                   "Median outcome (realistic)","80% chance EIRR is below this",
                   "90% chance EIRR is below this","Consultant's stated (optimistic)"]
    })
    st.dataframe(df_pct, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────────────
# TAB 2 — FRAGILITY DRIVERS
# ─────────────────────────────────────────────────────────────────────
with tab2:
    ct, cm = st.columns([3,2])
    with ct:
        st.markdown("#### Spearman Rank Tornado")
        st.caption("Rank correlation between each input variable and EIRR. Red = higher value → lower EIRR.")
        names=[t[0] for t in tornado[:7]][::-1]; rhos=[t[1] for t in tornado[:7]][::-1]
        fig_tor=go.Figure(go.Bar(x=rhos,y=names,orientation="h",
                                  marker_color=["#DC3545" if r<0 else "#0D6EFD" for r in rhos],
                                  opacity=0.85,text=[f"{r:+.3f}" for r in rhos],textposition="outside"))
        fig_tor.add_vline(x=0,line_color="#212529",line_width=1)
        fig_tor.update_layout(title=f"<b>Primary Driver: {tornado[0][0]}</b>  (ρ={tornado[0][1]:+.3f})",
                               height=380,plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                               xaxis=dict(title="Spearman ρ with EIRR",gridcolor="#EEEEEE"),
                               margin=dict(l=10,r=80,t=50,b=40),showlegend=False)
        st.plotly_chart(fig_tor, use_container_width=True)

    with cm:
        st.markdown("#### Procurement Mode Comparison")
        st.caption("Same project under 3 modes. Black outline = DPR's chosen mode.")
        with st.spinner("Computing all modes…"):
            all_fi_m={}
            for m in MODES:
                r_m=_sim(pj,m,min(n_iter,2000))[0]
                all_fi_m[m]=r_m["fi_p"]

        fig_mc=go.Figure(go.Bar(x=list(all_fi_m.keys()),y=list(all_fi_m.values()),
                                 marker_color=[_fc(f) for f in all_fi_m.values()],
                                 text=[f"{f:.0f}%" for f in all_fi_m.values()],
                                 textposition="outside",opacity=0.87,
                                 marker_line=dict(
                                     color=["rgba(0,0,0,1)" if m==p['dpr_mode'] else "rgba(255,255,255,0.5)" for m in MODES],
                                     width=[3 if m==p['dpr_mode'] else 0.5 for m in MODES])))
        fig_mc.add_hline(y=50,line_dash="dash",line_color="#DC3545",opacity=0.7)
        fig_mc.add_hline(y=25,line_dash="dash",line_color="#856404",opacity=0.7)
        fig_mc.add_hrect(y0=50,y1=110,fillcolor=RGBA["red_fill"],line_width=0)
        fig_mc.add_hrect(y0=25,y1=50, fillcolor=RGBA["amber_fill"],line_width=0)
        fig_mc.add_hrect(y0=0, y1=25, fillcolor=RGBA["green_fill"],line_width=0)
        fig_mc.update_layout(height=380,plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                              yaxis=dict(title="FI (%)",range=[0,115],gridcolor="#EEEEEE"),
                              margin=dict(l=40,r=60,t=30,b=40),showlegend=False)
        st.plotly_chart(fig_mc, use_container_width=True)
        best=min(all_fi_m,key=all_fi_m.get); worst=max(all_fi_m,key=all_fi_m.get)
        if all_fi_m[worst]-all_fi_m[best]>15:
            rec=f"DPR's **{p['dpr_mode']}** is optimal." if best==p['dpr_mode'] else \
                f"Consider **{best}** (FI={all_fi_m[best]:.0f}%) over **{p['dpr_mode']}** (FI={all_fi_m[p['dpr_mode']]:.0f}%). Δ={all_fi_m[worst]-all_fi_m[best]:.0f}pp"
            st.info(rec)


# ─────────────────────────────────────────────────────────────────────
# TAB 3 — SWITCHING VALUES (THE CORE FINDING)
# ─────────────────────────────────────────────────────────────────────
with tab3:
    st.markdown("#### Switching Values — Dual Anchor Analysis")
    st.markdown("""<div class='note'>
    <b>What is a Switching Value?</b> (UK Green Book §6.103): the minimum change in one variable,
    holding all others at their reference values, that makes the project stop representing value for money.<br><br>
    <b>Two anchors give two different answers:</b><br>
    • <b>DPR-Anchored (Consultant's View):</b> starts from DPR_EIRR. Shows how much tolerance the consultant claims.<br>
    • <b>P50-Anchored (PFFF's View):</b> starts from simulated P50. Shows actual tolerance under realistic uncertainty.<br><br>
    <b>For RED projects</b> (P50 < 12%): the project is <i>already below the viability threshold at median outcomes</i>.
    The DPR's "safe" switching values are phantom safety — the project cannot improve by only tolerating changes;
    it needs the consultant's assumptions to be correct. Delay SV shows "Already failed" — it is literally impossible
    to be "safe against delay" when P50 EIRR is already below 12%.
    </div>""", unsafe_allow_html=True)

    # The dual SV comparison cards
    st.markdown("##### Side-by-Side: Consultant's Claim vs PFFF Reality")
    col_dpr, col_div, col_p50 = st.columns([5,1,5])

    with col_dpr:
        st.markdown(f"""<div class='sv-dpr'>
        <b>🏦 DPR-Anchored (Consultant's View)</b><br>
        Starts from: <b>DPR EIRR = {p['dpr_eirr']:.2f}%</b><br>
        Headroom above 12%: <b>{svs['dpr_gap']:+.2f}pp</b><br><br>
        <b>Cost Overrun Tolerance:</b> {f'+{svs["dpr_cost"]:.1f}%' if svs['dpr_cost'] else '∞'}<br>
        <small>Consultant claims project survives {f'{svs["dpr_cost"]:.1f}%' if svs['dpr_cost'] else 'unlimited'} cost overrun</small><br><br>
        <b>Traffic Shortfall Tolerance:</b> {f'−{svs["dpr_traf"]:.1f}%' if svs['dpr_traf'] else '∞'}<br>
        <small>Consultant claims project survives {f'{svs["dpr_traf"]:.1f}%' if svs['dpr_traf'] else 'unlimited'} traffic shortfall</small><br><br>
        <b>Delay Tolerance:</b> {f'+{svs["dpr_delay"]:.0f} months' if svs['dpr_delay'] else '∞'}<br>
        <small>Consultant claims project survives {f'{svs["dpr_delay"]:.0f} months' if svs['dpr_delay'] else 'unlimited'} construction delay</small>
        </div>""", unsafe_allow_html=True)

    with col_div:
        st.markdown("<br><br><br><div style='text-align:center; font-size:2rem; color:#842029'>⟹</div>", unsafe_allow_html=True)

    with col_p50:
        if already_failed:
            st.markdown(f"""<div class='sv-p50'>
            <b>🔬 P50-Anchored (PFFF Reality)</b><br>
            Starts from: <b>P50 EIRR = {p50:.2f}%</b><br>
            Gap below 12%: <b>{svs['p50_gap']:+.2f}pp</b> ← <b>ALREADY FAILED AT MEDIAN</b><br><br>
            <b>Cost Overrun Tolerance:</b> {f'{svs["p50_cost"]:+.1f}%' if svs.get('p50_cost') else 'N/A'}<br>
            <small>{"Project needs cost to be " + str(abs(svs['p50_cost'])) + "% LOWER than DPR to achieve 12% at P50" if svs.get('p50_cost') and svs['p50_cost']<0 else "N/A"}</small><br><br>
            <b>Traffic Shortfall Tolerance:</b> Already below hurdle<br>
            <small>Any additional traffic shortfall makes it worse. The CAG average shortfall of 44% would push further into deficit.</small><br><br>
            <b>Delay Tolerance:</b> ⚠️ <b>None — project already below hurdle</b><br>
            <small>The DPR's delay SV of {svs['dpr_delay']:.0f}mo is phantom safety. At P50, project is {abs(svs['p50_gap']):.1f}pp below hurdle BEFORE any delay.</small>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class='sv-ok'>
            <b>🔬 P50-Anchored (PFFF Reality)</b><br>
            Starts from: <b>P50 EIRR = {p50:.2f}%</b><br>
            Headroom above 12%: <b>{svs['p50_gap']:+.2f}pp</b><br><br>
            <b>Cost Overrun Tolerance:</b> {f'+{svs["p50_cost"]:.1f}%' if svs.get('p50_cost') else '∞'}<br>
            <b>Traffic Shortfall Tolerance:</b> {f'−{svs["p50_traf"]:.1f}%' if svs.get('p50_traf') else '∞'}<br>
            <b>Delay Tolerance:</b> {f'+{svs["p50_delay"]:.0f}mo' if svs.get('p50_delay') else '∞'}<br>
            </div>""", unsafe_allow_html=True)

    # Bias summary
    st.markdown(f"""<div style='background:#F8D7DA;border-left:5px solid #842029;border-radius:8px;padding:14px;margin:12px 0;'>
    <b>The Bias Gap: {bias_gap:+.2f}pp</b><br>
    The consultant's DPR claimed {p['dpr_eirr']:.2f}% EIRR. PFFF's probabilistic simulation gives {p50:.2f}% at P50.
    The difference ({bias_gap:+.2f}pp) is the <b>Optimism Bias</b> captured by PFFF — the systematic overstatement
    that the current deterministic appraisal system cannot detect.<br>
    {'<br><b>CAG Reference:</b> Average Indian NH cost overrun = +71%. '
     + (f'The DPR cost SV of +{svs["dpr_cost"]:.1f}% is well below 71% — the project would already fail at average performance.' if svs.get("dpr_cost") and svs["dpr_cost"]<71 else
        f'The DPR cost SV of +{svs.get("dpr_cost","∞")} is above the 71% average — modest robustness on this metric.')
     if svs.get("dpr_cost") else ''}
    </div>""", unsafe_allow_html=True)

    # OAT charts
    st.divider()
    st.markdown("#### One-At-A-Time Sensitivity (OAT) Curves")
    st.caption("Starting from DPR values. Where curve crosses 12% = DPR-anchored SV. "
               "The gap between DPR_EIRR and the curve's starting point = optimism bias.")

    oc1,oc2=st.columns(2)
    with oc1:
        xr=np.linspace(-5,min(200,(svs.get('dpr_cost') or 100)*2+10),80)
        yr=[eirr_iter(p,scn,p["civil_cr"]*(1+x/100),0,p["yr1_aadt"],p["growth"],1.0,1.0)*100 for x in xr]
        fig_oc=go.Figure()
        fig_oc.add_trace(go.Scatter(x=xr,y=yr,mode="lines",line=dict(color="#DC3545",width=2.5)))
        fig_oc.add_hline(y=12,line_dash="dash",line_color="#212529",line_width=2,annotation_text="12% Hurdle")
        fig_oc.add_hline(y=p50,line_dash="dot",line_color="#0D6EFD",line_width=1.5,
                         annotation_text=f"P50={p50:.1f}%",annotation_position="right",annotation_font_color="#0D6EFD")
        if svs.get("dpr_cost"):
            fig_oc.add_vline(x=svs["dpr_cost"],line_dash="dot",line_color="#DC3545",line_width=1.5,
                             annotation_text=f"DPR-SV: +{svs['dpr_cost']:.1f}%",annotation_font_color="#DC3545")
        fig_oc.add_vline(x=71,line_dash="dot",line_color="#856404",line_width=1.5,
                         annotation_text="CAG avg: +71%",annotation_font_color="#856404")
        if already_failed:
            fig_oc.add_hrect(y0=min(yr),y1=12,fillcolor=RGBA["red_fill"],line_width=0)
        fig_oc.update_layout(title="<b>EIRR vs Civil Cost Overrun (OAT)</b>",height=360,
                              plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                              xaxis=dict(title="Cost Overrun (%)",gridcolor="#EEEEEE"),
                              yaxis=dict(title="EIRR (%)",gridcolor="#EEEEEE"),
                              margin=dict(l=50,r=50,t=50,b=40),showlegend=False)
        st.plotly_chart(fig_oc, use_container_width=True)

    with oc2:
        xr2=np.linspace(-5,min(90,(svs.get('dpr_traf') or 50)*2+10),80)
        yr2=[eirr_iter(p,scn,p["civil_cr"],0,p["yr1_aadt"]*(1-x/100),p["growth"],1.0,1.0)*100 for x in xr2]
        fig_ot=go.Figure()
        fig_ot.add_trace(go.Scatter(x=xr2,y=yr2,mode="lines",line=dict(color="#0D6EFD",width=2.5)))
        fig_ot.add_hline(y=12,line_dash="dash",line_color="#212529",line_width=2)
        fig_ot.add_hline(y=p50,line_dash="dot",line_color="#0D6EFD",line_width=1.5,
                         annotation_text=f"P50={p50:.1f}%",annotation_position="right",annotation_font_color="#0D6EFD")
        if svs.get("dpr_traf"):
            fig_ot.add_vline(x=svs["dpr_traf"],line_dash="dot",line_color="#0D6EFD",line_width=1.5,
                             annotation_text=f"DPR-SV: −{svs['dpr_traf']:.1f}%",annotation_font_color="#0D6EFD")
        fig_ot.add_vline(x=44,line_dash="dot",line_color="#856404",line_width=1.5,
                         annotation_text="Bain P10: 44%",annotation_font_color="#856404")
        if already_failed:
            fig_ot.add_hrect(y0=min(yr2),y1=12,fillcolor=RGBA["red_fill"],line_width=0)
        fig_ot.update_layout(title="<b>EIRR vs Traffic Shortfall (OAT)</b>",height=360,
                              plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                              xaxis=dict(title="Traffic Shortfall (%)",gridcolor="#EEEEEE"),
                              yaxis=dict(title="EIRR (%)",gridcolor="#EEEEEE"),
                              margin=dict(l=50,r=50,t=50,b=40),showlegend=False)
        st.plotly_chart(fig_ot, use_container_width=True)

    oc3,oc4=st.columns(2)
    with oc3:
        dm=min(150,(svs.get('dpr_delay') or 72)*1.5+12)
        dr=np.linspace(0,dm,60)
        yd=[eirr_iter(p,scn,p["civil_cr"],d,p["yr1_aadt"],p["growth"],1.0,1.0)*100 for d in dr]
        fig_od=go.Figure()
        fig_od.add_trace(go.Scatter(x=dr,y=yd,mode="lines",line=dict(color="#6F42C1",width=2.5)))
        fig_od.add_hline(y=12,line_dash="dash",line_color="#212529",line_width=2)
        fig_od.add_hline(y=p50,line_dash="dot",line_color="#0D6EFD",line_width=1.5,
                         annotation_text=f"P50={p50:.1f}%",annotation_position="right",annotation_font_color="#0D6EFD")
        if svs.get("dpr_delay"):
            fig_od.add_vline(x=svs["dpr_delay"],line_dash="dot",line_color="#6F42C1",line_width=1.5,
                             annotation_text=f"DPR-SV: {svs['dpr_delay']:.0f}mo",annotation_font_color="#6F42C1")
        if already_failed:
            fig_od.add_hrect(y0=min(yd),y1=12,fillcolor=RGBA["red_fill"],line_width=0)
            fig_od.add_annotation(x=dm*0.5,y=(min(yd)+12)/2,text="P50 already below 12%\nAny delay makes it worse",
                                   showarrow=False,font=dict(color="#842029",size=10),bgcolor="white",
                                   bordercolor="#842029",borderwidth=1.5,borderpad=4)
        fig_od.update_layout(title="<b>EIRR vs Construction Delay (OAT)</b>",height=360,
                              plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                              xaxis=dict(title="Delay (months)",gridcolor="#EEEEEE"),
                              yaxis=dict(title="EIRR (%)",gridcolor="#EEEEEE"),
                              margin=dict(l=50,r=50,t=50,b=40),showlegend=False)
        st.plotly_chart(fig_od, use_container_width=True)

    with oc4:
        gr=np.linspace(0.01,0.13,60)
        yg=[eirr_iter(p,scn,p["civil_cr"],0,p["yr1_aadt"],g,1.0,1.0)*100 for g in gr]
        fig_og=go.Figure()
        fig_og.add_trace(go.Scatter(x=gr*100,y=yg,mode="lines",line=dict(color="#198754",width=2.5)))
        fig_og.add_hline(y=12,line_dash="dash",line_color="#212529",line_width=2)
        fig_og.add_hline(y=p50,line_dash="dot",line_color="#0D6EFD",line_width=1.5,
                         annotation_text=f"P50={p50:.1f}%",annotation_position="right",annotation_font_color="#0D6EFD")
        fig_og.add_vline(x=p["growth"]*100,line_dash="dot",line_color="#198754",line_width=1.5,
                         annotation_text=f"DPR: {p['growth']*100:.1f}%",annotation_font_color="#198754")
        fig_og.update_layout(title="<b>EIRR vs Traffic Growth Rate (OAT)</b>",height=360,
                              plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                              xaxis=dict(title="Growth Rate (% p.a.)",gridcolor="#EEEEEE"),
                              yaxis=dict(title="EIRR (%)",gridcolor="#EEEEEE"),
                              margin=dict(l=50,r=50,t=50,b=40),showlegend=False)
        st.plotly_chart(fig_og, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────
# TAB 4 — ALL 7 PROJECTS
# ─────────────────────────────────────────────────────────────────────
with tab4:
    st.markdown("#### All 7 Projects — FI + Optimism Bias + Switching Values")
    with st.spinner("Running all 7 projects…"):
        batch={}
        for code,proj in PROJECTS.items():
            pj2=json.dumps(proj,default=str)
            r2=_sim(pj2,proj["dpr_mode"],min(n_iter,2000))
            batch[code]={"res":r2[0],"scn":r2[1],"svs":r2[5],"p50":r2[6]}

    codes_b=list(PROJECTS.keys())
    fis_b=[batch[c]["res"]["fi_p"] for c in codes_b]
    p50s_b=[batch[c]["p50"] for c in codes_b]
    dpr_b=[PROJECTS[c]["dpr_eirr"] for c in codes_b]
    bias_b=[d-p for d,p in zip(dpr_b,p50s_b)]

    fig_bat=go.Figure(go.Bar(
        x=codes_b,y=fis_b,marker_color=[_fc(f) for f in fis_b],opacity=0.87,
        text=[f"{f:.0f}%" for f in fis_b],textposition="outside",
        hovertemplate="%{x}: FI=%{y:.1f}%<extra></extra>"))
    fig_bat.add_hline(y=50,line_dash="dash",line_color="#DC3545",opacity=0.7,annotation_text="RED 50%")
    fig_bat.add_hline(y=25,line_dash="dash",line_color="#856404",opacity=0.7,annotation_text="AMBER 25%")
    fig_bat.add_hrect(y0=50,y1=110,fillcolor=RGBA["red_fill"],line_width=0)
    fig_bat.add_hrect(y0=25,y1=50, fillcolor=RGBA["amber_fill"],line_width=0)
    fig_bat.add_hrect(y0=0, y1=25, fillcolor=RGBA["green_fill"],line_width=0)
    for c in codes_b:
        if PROJECTS[c]["role"]=="VALIDATION":
            fig_bat.add_annotation(x=c,y=102,text="VALIDATION",showarrow=False,
                                   font=dict(size=8,color="#6c757d"),yanchor="bottom")
    fig_bat.update_layout(height=380,plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                          yaxis=dict(title="Fragility Index FI%",range=[0,115],gridcolor="#EEEEEE"),
                          margin=dict(l=50,r=50,t=30,b=40),showlegend=False)
    st.plotly_chart(fig_bat, use_container_width=True)

    # Bias chart
    x2=list(range(len(codes_b)))
    fig_bias=go.Figure()
    fig_bias.add_trace(go.Bar(x=codes_b,y=dpr_b,name="DPR EIRR (Consultant)",
                               marker_color="rgba(44,62,80,0.85)",width=0.35,
                               text=[f"{v:.1f}%" for v in dpr_b],textposition="outside"))
    fig_bias.add_trace(go.Bar(x=codes_b,y=p50s_b,name="P50 Simulated (PFFF)",
                               marker_color=[_fc(f) for f in fis_b],width=0.35,
                               text=[f"{v:.1f}%" for v in p50s_b],textposition="outside",
                               base=None))
    fig_bias.add_hline(y=12,line_dash="dash",line_color="#DC3545",line_width=2,annotation_text="12% Hurdle")
    fig_bias.update_layout(barmode="group",title="<b>Optimism Bias: DPR EIRR vs P50 Simulated</b><br>"
                           "<sup>Dark = Consultant's claim | Coloured = PFFF realistic P50 | Gap = Optimism Bias</sup>",
                           height=380,plot_bgcolor="#FAFAFA",paper_bgcolor="white",
                           yaxis=dict(title="EIRR (%)",gridcolor="#EEEEEE"),
                           legend=dict(orientation="h",y=1.08),
                           margin=dict(l=50,r=50,t=80,b=40))
    st.plotly_chart(fig_bias, use_container_width=True)

    # Summary table
    rows=[]
    for c in codes_b:
        sv=batch[c]["svs"]; p50_=batch[c]["p50"]
        rows.append({
            "Code":c,"Project":PROJECTS[c]["short"],"Mode":PROJECTS[c]["dpr_mode"],
            "DPR EIRR (%)":PROJECTS[c]["dpr_eirr"],
            "P50 EIRR (%)":round(p50_,2),
            "Bias (pp)":round(PROJECTS[c]["dpr_eirr"]-p50_,2),
            "FI Primary (%)":round(batch[c]["res"]["fi_p"],1),
            "Cost SV (DPR)":f"+{sv['dpr_cost']:.0f}%" if sv.get("dpr_cost") else "∞",
            "Cost SV (P50)":f"{sv['p50_cost']:+.0f}%" if sv.get("p50_cost") else "Deficit",
            "P50 Status":sv["p50_status"],
            "Verdict":[_vt(batch[c]["res"]["fi_p"])],
        })
    df_b=pd.DataFrame(rows); df_b=df_b.drop(columns=["Verdict"])
    st.dataframe(df_b, use_container_width=True, hide_index=True)

    p5fi=batch["P5"]["res"]["fi_p"]; p7fi=batch["P7"]["res"]["fi_p"]
    st.markdown(f"""<div class='note' style='border-left-color:{"#198754" if p5fi>=50 and p7fi>=25 else "#856404"}'>
    <b>Validation:</b> P5 Vadodara FI = <b>{p5fi:.1f}%</b>
    {"(RED ✓ — defaulted in reality)" if p5fi>=50 else "(⚠️ Expected RED)"} |
    P7 Samruddhi FI = <b>{p7fi:.1f}%</b>
    {"(AMBER-RED ✓ — traffic beat rescued a fragile project)" if p7fi>=25 else "(⚠️ Expected AMBER-RED)"}
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────
# TAB 5 — EXPORT
# ─────────────────────────────────────────────────────────────────────
with tab5:
    st.markdown("#### Export Audit Data")

    def _build_excel():
        try:
            from openpyxl import Workbook as WB
            from openpyxl.styles import PatternFill as PF, Font as FN, Alignment as AL
            from openpyxl.utils import get_column_letter as gcl
        except ImportError: return None

        wb=WB(); n_=len(samp["v05"])
        ws1=wb.active; ws1.title="Iterations"
        hdrs=["Iter","EIRR_%","FIRR_%","Equity_%","Civil_Cr","LA_Cr","Delay_Mo",
              "AADT","Growth_%","VOC","VoT","Stall"]
        for j,h in enumerate(hdrs,1):
            c=ws1.cell(1,j); c.value=h; c.font=FN(bold=True,color="FFFFFF")
            c.fill=PF("solid",fgColor="1F497D"); c.alignment=AL(horizontal="center")
        for i in range(n_):
            ws1.cell(i+2,1).value=i+1
            ws1.cell(i+2,2).value=round(res["eirr_arr"][i]*100,4)
            ws1.cell(i+2,3).value=round(res["firr_arr"][i]*100,4) if not np.isnan(res["firr_arr"][i]) else "N/A"
            ws1.cell(i+2,4).value=round(res["eq_arr"][i]*100,4) if not np.isnan(res["eq_arr"][i]) else "N/A"
            ws1.cell(i+2,5).value=round(samp["v05"][i],2); ws1.cell(i+2,6).value=round(samp["v06"][i],2)
            ws1.cell(i+2,7).value=round(samp["v07"][i],2); ws1.cell(i+2,8).value=round(samp["v01"][i],0)
            ws1.cell(i+2,9).value=round(samp["v02"][i]*100,4)
            ws1.cell(i+2,10).value=round(samp["v10"][i],4); ws1.cell(i+2,11).value=round(samp["v11"][i],4)
            ws1.cell(i+2,12).value=int(samp["reg"][i])
        for j in range(1,13): ws1.column_dimensions[gcl(j)].width=14

        ws2=wb.create_sheet("Audit Summary")
        zs_v,_=_zs_proof(pj)
        rows2=[("Project",p["name"]),("Mode Simulated",sim_mode),("DPR Mode",p["dpr_mode"]),
               ("Iterations",n_),("DPR EIRR (%)",p["dpr_eirr"]),("Zero-Stress EIRR (%)",round(zs_v,4)),
               ("FI Primary (%)",round(fi,2)),("FI EIRR (%)",round(res["fi_eirr"],2)),
               ("FI FIRR (%)",round(res["fi_firr"],2) if not np.isnan(res["fi_firr"]) else "N/A"),
               ("FI Equity (%)",round(res["fi_eq"],2) if not np.isnan(res["fi_eq"]) else "N/A"),
               ("P50 EIRR (%)",round(p50,2)),("Bias (DPR−P50, pp)",round(bias_gap,2)),
               ("P50 Status",svs["p50_status"]),
               ("Cost SV DPR-anchored",f"+{svs['dpr_cost']:.1f}%" if svs['dpr_cost'] else "∞"),
               ("Cost SV P50-anchored",f"{svs['p50_cost']:+.1f}%" if svs.get('p50_cost') else "Deficit"),
               ("Traffic SV DPR",f"-{svs['dpr_traf']:.1f}%" if svs['dpr_traf'] else "∞"),
               ("Delay SV DPR",f"+{svs['dpr_delay']:.0f}mo" if svs['dpr_delay'] else "∞"),
               ("Delay SV P50","Already below hurdle" if already_failed else f"+{svs.get('p50_delay',0):.0f}mo"),
               ("Primary Driver",tornado[0][0] if tornado else "—"),
               ("Verdict",_vt(fi))]
        for i,(k,v) in enumerate(rows2,1):
            ws2.cell(i,1).value=k; ws2.cell(i,1).font=FN(bold=True); ws2.cell(i,2).value=v
        ws2.column_dimensions["A"].width=35; ws2.column_dimensions["B"].width=30

        ws3=wb.create_sheet("Fragility Drivers")
        for j,h in enumerate(["Variable","Spearman ρ","Direction"],1):
            ws3.cell(1,j).value=h; ws3.cell(1,j).font=FN(bold=True)
        for i,(nm,rho) in enumerate(tornado,2):
            ws3.cell(i,1).value=nm; ws3.cell(i,2).value=round(rho,4)
            ws3.cell(i,3).value="Higher → lower EIRR" if rho<0 else "Higher → higher EIRR"
        for c_ in ["A","B","C"]: ws3.column_dimensions[c_].width=28
        return wb

    c_e1,c_e2=st.columns(2)
    with c_e1:
        if st.button("📊 Generate Excel Report",type="primary",use_container_width=True):
            with st.spinner("Building…"):
                wb_out=_build_excel()
            if wb_out:
                buf=io.BytesIO(); wb_out.save(buf)
                st.download_button("⬇️ Download Excel (3 sheets)",data=buf.getvalue(),
                                   file_name=f"PFFF_{p['name'][:20].replace(' ','_')}.xlsx",
                                   mime="application/vnd.ms-excel",use_container_width=True)
            else: st.error("openpyxl not available")

    with c_e2:
        df_csv=pd.DataFrame({"EIRR_%":res["eirr_arr"]*100,"FIRR_%":res["firr_arr"]*100,
                              "Equity_%":res["eq_arr"]*100,"Civil_Cr":samp["v05"],
                              "LA_Cr":samp["v06"],"Delay_Mo":samp["v07"],"AADT":samp["v01"]})
        st.download_button("⬇️ Download CSV (iterations)",data=df_csv.to_csv(index=False),
                           file_name=f"PFFF_{p['name'][:15].replace(' ','_')}.csv",
                           mime="text/csv",use_container_width=True)

st.divider()
st.caption("PFFF v13 | M.BEM Thesis 2024 | SPA Delhi | Varshni M S | Supervisor: Mr. Rhijul Sood | "
           "IRC SP:30:2019 · CAG 19/2023 · LARR 2013 · Flyvbjerg 2003 · Bain 2009 · UK Green Book 2022")
