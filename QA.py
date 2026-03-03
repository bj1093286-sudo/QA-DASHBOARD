# ============================================================
# 교육품질 통합 대시보드 v5.0
# save as: app.py
# run:     streamlit run app.py
# install: pip install streamlit plotly pandas numpy openpyxl
# ============================================================

import re, io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ─────────────────────────────────────────
# 0. PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="교육품질 통합 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# 1. CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
.stApp { background:#f4f6fb; }
.kpi-card {
    background:#fff; border-radius:12px;
    padding:16px 18px; box-shadow:0 2px 10px rgba(30,40,90,.10);
    text-align:center; margin-bottom:8px;
}
.kpi-label { font-size:.72rem; color:#6b7a99; font-weight:700;
             letter-spacing:.5px; text-transform:uppercase; }
.kpi-value { font-size:1.9rem; font-weight:800; color:#1a237e; line-height:1.2; margin:4px 0; }
.kpi-sub   { font-size:.75rem; color:#888; }
.kpi-up    { color:#1b8a4c; font-size:.8rem; }
.kpi-down  { color:#c0392b; font-size:.8rem; }
.kpi-eq    { color:#888;    font-size:.8rem; }
.sec-hd {
    background:linear-gradient(90deg,#1a237e,#3949ab);
    color:#fff; border-radius:8px;
    padding:8px 16px; font-size:.97rem;
    font-weight:700; margin:18px 0 10px;
}
.alert-box  { background:#fff3cd; border-left:4px solid #ffc107;
              border-radius:6px; padding:9px 14px; margin:6px 0; font-size:.87rem; }
.danger-box { background:#fdecea; border-left:4px solid #ef5350;
              border-radius:6px; padding:9px 14px; margin:6px 0; font-size:.87rem; }
.good-box   { background:#e8f5e9; border-left:4px solid #43a047;
              border-radius:6px; padding:9px 14px; margin:6px 0; font-size:.87rem; }
div[data-baseweb="tab"] { font-weight:700; padding:8px 20px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 2. HELPERS
# ─────────────────────────────────────────
COLORS = ["#1a237e","#3949ab","#e53935","#ffa726",
          "#42a5f5","#26c6da","#66bb6a","#ab47bc","#ec407a"]

CHART_BASE = dict(
    plot_bgcolor="#fff", paper_bgcolor="#fff",
    font=dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif", size=12),
    margin=dict(t=50,b=40,l=40,r=20),
)

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

def safe_mean(s):
    v = to_num(s).dropna()
    return float(v.mean()) if len(v) else 0.0

def clean_col(c):
    return re.sub(r"\s+","",str(c)).strip()

def kpi(label, value, sub=None, delta=None, fmt=".1f", suffix=""):
    v = f"{value:{fmt}}{suffix}" if isinstance(value,(int,float)) else str(value)
    sh = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    if delta is None: d=""
    elif delta>0:  d=f'<div class="kpi-up">▲ {delta:+.1f}</div>'
    elif delta<0:  d=f'<div class="kpi-down">▼ {delta:+.1f}</div>'
    else:          d='<div class="kpi-eq">― 변동없음</div>'
    return f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{v}</div>{sh}{d}</div>'

def sec(txt):
    st.markdown(f'<div class="sec-hd">{txt}</div>', unsafe_allow_html=True)

def box(txt, kind="alert"):
    cls={"alert":"alert-box","danger":"danger-box","good":"good-box"}.get(kind,"alert-box")
    st.markdown(f'<div class="{cls}">{txt}</div>', unsafe_allow_html=True)

def wk_sort_key(w):
    try:
        m=int(re.search(r"(\d+)월",str(w)).group(1))
        n=int(re.search(r"(\d+)[주W]",str(w)).group(1))
        return m*10+n
    except: return 0

def parse_tsv(text: str) -> pd.DataFrame:
    """붙여넣기 텍스트(탭/공백 구분) → DataFrame"""
    text = text.strip()
    if not text:
        return pd.DataFrame()
    try:
        sep = "\t" if "\t" in text else ","
        df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str)
        df.columns = [clean_col(c) for c in df.columns]
        return df.dropna(how="all")
    except:
        return pd.DataFrame()

# ─────────────────────────────────────────
# 3. SIDEBAR
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 교육품질 대시보드")
    st.markdown("---")
    report_ym    = st.text_input("📅 보고 기준 월", value="2026년 1월")
    prev_ym      = st.text_input("📅 전월", value="2025년 12월")
    qa_target    = st.number_input("🎯 QA 목표점수",   value=90.0, step=0.5)
    test_pass    = st.number_input("📝 직무테스트 기준", value=80.0, step=1.0)
    csat_target  = st.number_input("⭐ CSAT 목표점수",  value=92.0, step=0.5)
    csat_monitor = st.number_input("🔴 CSAT 모니터링 기준", value=70.0, step=1.0)
    st.markdown("---")
    st.caption("데이터를 각 탭에 붙여넣기(Ctrl+V)하세요")

# ─────────────────────────────────────────
# 4. MAIN TITLE
# ─────────────────────────────────────────
st.markdown(f"""
<div style="background:linear-gradient(90deg,#1a237e,#3949ab);
     color:#fff;border-radius:12px;padding:22px 28px;margin-bottom:20px;">
  <h2 style="margin:0;font-size:1.55rem;">📊 교육품질 통합 대시보드</h2>
  <p style="margin:4px 0 0;opacity:.85;font-size:.92rem;">
    {report_ym} 보고 · QA목표 {qa_target}점 · 직무기준 {test_pass}점 · CSAT목표 {csat_target}점
  </p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 5. TABS
# ─────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📞 QA (Call)",
    "💬 QA (Chat)",
    "📝 직무테스트",
    "⭐ CSAT",
])

# ══════════════════════════════════════════════════
# TAB1 : QA CALL
# ══════════════════════════════════════════════════
with tab1:
    st.markdown("### 📞 정기 평가 결과 (Call)")

    with st.expander("📋 데이터 입력 방법", expanded=False):
        st.markdown("""
        **엑셀에서 데이터 복사 → 아래 텍스트 박스에 붙여넣기(Ctrl+V)**
        
        필요 컬럼: `구분(월/주차)`, `근속그룹`, `평균`, `1주차`, `2주차`, `3주차`, `인원`  
        항목별 이행률: `구분`, `평가건수`, `첫인사`, `정보확인`, `끝인사`, `인사톤`, `통화종료`,  
        `음성숙련도`, `전반적인감정연출`, `양해`, `즉각호응`, `언어표현`, `경청`, `문의파악`,  
        `맞춤설명`, `정확한안내`, `프로세스`, `전산처리`, `상담이력`, `가점`, `감점`
        """)

    c_left, c_right = st.columns(2)

    with c_left:
        st.markdown("**📊 근속그룹별 주차별 점수 (붙여넣기)**")
        qa_call_score_txt = st.text_area(
            "근속그룹별 점수 데이터",
            height=160,
            placeholder="근속그룹\t평균\t1주차\t2주차\t3주차\t인원\n신입3(90일이내)\t90.9\t92.7\t92.1\t87.7\t1",
            key="qa_call_score",
            label_visibility="collapsed",
        )

    with c_right:
        st.markdown("**📋 전월 비교 데이터 (선택)**")
        qa_call_prev_txt = st.text_area(
            "전월 데이터",
            height=160,
            placeholder="근속그룹\t전월평균\n신입3(90일이내)\t91.3",
            key="qa_call_prev",
            label_visibility="collapsed",
        )

    st.markdown("**📋 항목별 이행률 (붙여넣기)**")
    qa_call_item_txt = st.text_area(
        "항목별 이행률",
        height=160,
        placeholder="구분\t평가건수\t첫인사\t정보확인\t끝인사\t인사톤\t통화종료\t음성숙련도\t전반적인감정연출\t양해\t즉각호응\t언어표현\t경청\t문의파악\t맞춤설명\t정확한안내\t프로세스\t전산처리\t상담이력\t가점\t감점\n01월\t41\t97.6%\t95.1%\t97.6%\t100.0%\t100.0%\t92.7%\t90.2%\t87.8%\t95.1%\t99.0%\t82.9%\t95.1%\t100.0%\t87.8%\t90.2%\t80.5%\t100.0%\t76.8%\t0.0%\t0.0%",
        key="qa_call_item",
        label_visibility="collapsed",
    )

    if st.button("📊 Call QA 장표 생성", key="btn_call", type="primary"):
        # ── 점수 데이터 파싱 ──
        score_df = parse_tsv(qa_call_score_txt)
        item_df  = parse_tsv(qa_call_item_txt)
        prev_df  = parse_tsv(qa_call_prev_txt)

        if score_df.empty and item_df.empty:
            box("데이터를 붙여넣기 해주세요.", "alert")
        else:
            # ── KPI ──
            sec("① KPI 요약")
            avg_col = next((c for c in score_df.columns if "평균" in c), None)
            인원_col = next((c for c in score_df.columns if "인원" in c), None)

            if avg_col and not score_df.empty:
                # Total 행 분리
                total_row = score_df[score_df.iloc[:,0].str.contains("otal|합계|전체|Total", na=False, case=False)]
                group_row = score_df[~score_df.iloc[:,0].str.contains("otal|합계|전체|Total|Gap|gap", na=False, case=False)]

                curr_avg = safe_mean(score_df[avg_col]) if total_row.empty else safe_mean(total_row[avg_col])
                total_cnt = safe_mean(score_df[인원_col]) if 인원_col else 0

                prev_avg = 0.0
                if not prev_df.empty and avg_col in prev_df.columns:
                    prev_avg = safe_mean(prev_df[avg_col])
                delta_v = curr_avg - prev_avg if prev_avg else None

                cols_kpi = st.columns(5)
                cols_kpi[0].markdown(kpi("당월 평균점수", curr_avg, sub=f"목표 {qa_target}점", delta=delta_v), unsafe_allow_html=True)
                cols_kpi[1].markdown(kpi("전월 평균점수", prev_avg if prev_avg else "-", fmt=".1f"), unsafe_allow_html=True)
                cols_kpi[2].markdown(kpi("총 평가인원", int(total_cnt) if total_cnt else len(score_df), fmt=",d", sub="명"), unsafe_allow_html=True)
                cols_kpi[3].markdown(kpi("목표점수", qa_target, fmt=".1f"), unsafe_allow_html=True)
                gap_v = curr_avg - qa_target
                cols_kpi[4].markdown(kpi("목표 GAP", gap_v, fmt="+.1f"), unsafe_allow_html=True)

                if curr_avg >= qa_target:
                    box(f"✅ 목표점수 {qa_target}점 달성! 현재 <b>{curr_avg:.1f}점</b>", "good")
                else:
                    box(f"⚠️ 목표점수 {qa_target}점 미달성. 현재 <b>{curr_avg:.1f}점</b> (GAP {gap_v:+.1f}점)", "danger")

            # ── 차트 ──
            st.markdown("---")
            c1, c2 = st.columns([3, 2])

            with c1:
                sec("② 근속그룹별 주차별 점수 추이")
                if not score_df.empty:
                    wk_cols = [c for c in score_df.columns if "주차" in c or "WK" in c.upper()]
                    group_col = score_df.columns[0]

                    # 평균 열 + 주차열 long format
                    plot_cols = ([avg_col] if avg_col else []) + wk_cols
                    if plot_cols:
                        melt_df = score_df[~score_df[group_col].str.contains("Gap|gap|GAP", na=False)].copy()
                        melt_df = melt_df[[group_col] + plot_cols].melt(
                            id_vars=group_col, var_name="구분", value_name="점수"
                        )
                        melt_df["점수"] = to_num(melt_df["점수"])

                        fig = px.bar(
                            melt_df[melt_df["구분"].isin(wk_cols)] if wk_cols else melt_df,
                            x=group_col, y="점수", color="구분",
                            barmode="group",
                            text="점수",
                            color_discrete_sequence=COLORS,
                            title=f"근속그룹별 주차별 점수 ({report_ym})",
                        )
                        # 평균선 오버레이
                        if avg_col and avg_col in score_df.columns:
                            avg_map = score_df.set_index(group_col)[avg_col].apply(to_num)
                            fig.add_trace(go.Scatter(
                                x=score_df[~score_df[group_col].str.contains("Gap|gap",na=False)][group_col],
                                y=avg_map.dropna().values,
                                mode="lines+markers+text",
                                name="평균",
                                line=dict(color="#ffa726", width=3),
                                marker=dict(size=8),
                                text=avg_map.dropna().round(1).values,
                                textposition="top center",
                            ))
                        fig.add_hline(y=qa_target, line_dash="dot", line_color="#ef5350",
                                      annotation_text=f"목표 {qa_target}점")
                        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside",
                                          selector=dict(type="bar"))
                        fig.update_layout(**CHART_BASE, height=380,
                                          yaxis=dict(range=[75,105]),
                                          legend=dict(orientation="h", y=1.08))
                        st.plotly_chart(fig, use_container_width=True)

            with c2:
                sec("③ 근속그룹별 주차별 결과표")
                if not score_df.empty:
                    disp = score_df.copy()
                    st.dataframe(disp, use_container_width=True, hide_index=True, height=300)

            # ── 항목별 이행률 ──
            sec("④ 항목별 이행률")
            if not item_df.empty:
                구분_col = item_df.columns[0]
                num_cols = [c for c in item_df.columns if c not in [구분_col, "평가건수"]]

                # % 문자 제거 후 숫자 변환
                item_num = item_df.copy()
                for c in num_cols:
                    item_num[c] = to_num(item_num[c].astype(str).str.replace("%",""))

                # 히트맵
                gap_row   = item_num[item_num[구분_col].str.contains("Gap|gap|GAP", na=False)]
                data_rows = item_num[~item_num[구분_col].str.contains("Gap|gap|GAP", na=False)]

                if not data_rows.empty and num_cols:
                    z_vals = data_rows[num_cols].values.tolist()
                    fig2 = go.Figure(go.Heatmap(
                        z=z_vals,
                        x=num_cols,
                        y=data_rows[구분_col].tolist(),
                        colorscale="Blues",
                        zmin=60, zmax=100,
                        text=[[f"{v:.1f}%" if not pd.isna(v) else "-" for v in row] for row in z_vals],
                        texttemplate="%{text}",
                        textfont={"size":10},
                    ))
                    fig2.update_layout(**CHART_BASE, title="항목별 이행률 히트맵",
                                       height=220,
                                       xaxis=dict(tickangle=-30))
                    st.plotly_chart(fig2, use_container_width=True)

                # 원본 표
                st.dataframe(item_df, use_container_width=True, hide_index=True)

                # GAP 강조
                if not gap_row.empty:
                    sec("⑤ GAP 분석 (전월 대비)")
                    gap_melt = gap_row.melt(id_vars=구분_col, var_name="항목", value_name="GAP")
                    gap_melt["GAP_num"] = to_num(
                        gap_melt["GAP"].astype(str).str.replace("%","").str.replace("p","")
                    )
                    gap_melt = gap_melt.dropna(subset=["GAP_num"])
                    if not gap_melt.empty:
                        colors_gap = ["#ef5350" if v < 0 else "#43a047" for v in gap_melt["GAP_num"]]
                        fig3 = go.Figure(go.Bar(
                            x=gap_melt["항목"], y=gap_melt["GAP_num"],
                            marker_color=colors_gap,
                            text=gap_melt["GAP"].values,
                            textposition="outside",
                        ))
                        fig3.add_hline(y=0, line_color="#333", line_width=1)
                        fig3.update_layout(**CHART_BASE, height=300,
                                           title="전월 대비 GAP",
                                           yaxis_title="GAP(p)")
                        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    st.text_area("💬 Comment", key="call_cmt",
                 placeholder="Call QA 특이사항을 입력하세요.")
    st.text_area("📋 02월 목표 및 계획", key="call_plan",
                 placeholder="목표 및 Action Plan을 입력하세요.")

# ══════════════════════════════════════════════════
# TAB2 : QA CHAT
# ══════════════════════════════════════════════════
with tab2:
    st.markdown("### 💬 정기 평가 결과 (Chat)")

    with st.expander("📋 데이터 입력 방법", expanded=False):
        st.markdown("""
        **필요 컬럼:** `상담사`, `평균`, `1주차`, `2주차`, `3주차`  
        **항목별 이행률:** Call과 동일 컬럼 구조
        """)

    c_left2, c_right2 = st.columns(2)
    with c_left2:
        st.markdown("**📊 상담사별 주차별 점수**")
        qa_chat_score_txt = st.text_area(
            "상담사별 점수",
            height=160,
            placeholder="상담사\t평균\t1주차\t2주차\t3주차\n문채희\t97.5\t95.0\t100.0\t-",
            key="qa_chat_score",
            label_visibility="collapsed",
        )
    with c_right2:
        st.markdown("**📋 전월 비교 데이터 (선택)**")
        qa_chat_prev_txt = st.text_area(
            "전월",
            height=160,
            placeholder="상담사\t전월평균\n문채희\t88.8",
            key="qa_chat_prev",
            label_visibility="collapsed",
        )

    st.markdown("**📋 항목별 이행률**")
    qa_chat_item_txt = st.text_area(
        "항목별 이행률",
        height=160,
        placeholder="구분\t평가건수\t첫인사\t정보확인\t끝인사\t양해\t즉각호응\t대기\t언어표현\t가독성\t문의파악\t맞춤설명\t정확한안내\t프로세스\t전산처리\t상담이력\t가점\t감점",
        key="qa_chat_item",
        label_visibility="collapsed",
    )

    if st.button("📊 Chat QA 장표 생성", key="btn_chat", type="primary"):
        score_df2 = parse_tsv(qa_chat_score_txt)
        item_df2  = parse_tsv(qa_chat_item_txt)
        prev_df2  = parse_tsv(qa_chat_prev_txt)

        if score_df2.empty and item_df2.empty:
            box("데이터를 붙여넣기 해주세요.", "alert")
        else:
            avg_col2 = next((c for c in score_df2.columns if "평균" in c), None)

            # KPI
            sec("① KPI 요약")
            curr_avg2 = safe_mean(score_df2[avg_col2]) if avg_col2 else 0
            prev_avg2 = 0.0
            if not prev_df2.empty and avg_col2 in prev_df2.columns:
                prev_avg2 = safe_mean(prev_df2[avg_col2])
            delta2 = curr_avg2 - prev_avg2 if prev_avg2 else None

            cols_k2 = st.columns(4)
            cols_k2[0].markdown(kpi("당월 평균점수", curr_avg2, delta=delta2), unsafe_allow_html=True)
            cols_k2[1].markdown(kpi("전월 평균점수", prev_avg2 if prev_avg2 else "-", fmt=".1f"), unsafe_allow_html=True)
            cols_k2[2].markdown(kpi("목표점수", qa_target, fmt=".1f"), unsafe_allow_html=True)
            cols_k2[3].markdown(kpi("GAP", curr_avg2 - qa_target, fmt="+.1f"), unsafe_allow_html=True)

            st.markdown("---")
            c1c, c2c = st.columns([3,2])

            with c1c:
                sec("② 상담사별 주차별 점수 추이")
                if not score_df2.empty:
                    agent_col = score_df2.columns[0]
                    wk_cols2  = [c for c in score_df2.columns if "주차" in c]
                    plot_df2  = score_df2[~score_df2[agent_col].str.contains(
                        "otal|합계|Gap|gap|전체",na=False,case=False)].copy()

                    if wk_cols2:
                        melt2 = plot_df2[[agent_col]+wk_cols2].melt(
                            id_vars=agent_col, var_name="주차", value_name="점수")
                        melt2["점수"] = to_num(melt2["점수"])
                        fig4 = px.line(
                            melt2, x="주차", y="점수", color=agent_col,
                            markers=True, text="점수",
                            color_discrete_sequence=COLORS,
                            title=f"상담사별 주차별 점수 ({report_ym})",
                        )
                        fig4.add_hline(y=qa_target, line_dash="dot", line_color="#ef5350",
                                       annotation_text=f"목표 {qa_target}점")
                        fig4.update_traces(texttemplate="%{text:.1f}", textposition="top center")
                        fig4.update_layout(**CHART_BASE, height=380,
                                           yaxis=dict(range=[75,105]),
                                           legend=dict(orientation="h", y=1.08))
                        st.plotly_chart(fig4, use_container_width=True)

            with c2c:
                sec("③ 상담사별 결과표")
                st.dataframe(score_df2, use_container_width=True, hide_index=True, height=300)

            # 항목별 이행률
            sec("④ 항목별 이행률")
            if not item_df2.empty:
                구분2 = item_df2.columns[0]
                num2  = [c for c in item_df2.columns if c not in [구분2,"평가건수"]]
                item_num2 = item_df2.copy()
                for c in num2:
                    item_num2[c] = to_num(item_num2[c].astype(str).str.replace("%",""))

                data2 = item_num2[~item_num2[구분2].str.contains("Gap|gap|GAP",na=False)]
                if not data2.empty and num2:
                    z2 = data2[num2].values.tolist()
                    fig5 = go.Figure(go.Heatmap(
                        z=z2, x=num2, y=data2[구분2].tolist(),
                        colorscale="Blues", zmin=60, zmax=100,
                        text=[[f"{v:.1f}%" if not pd.isna(v) else "-" for v in row] for row in z2],
                        texttemplate="%{text}", textfont={"size":10},
                    ))
                    fig5.update_layout(**CHART_BASE, title="항목별 이행률 히트맵",
                                       height=220, xaxis=dict(tickangle=-30))
                    st.plotly_chart(fig5, use_container_width=True)

                st.dataframe(item_df2, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.text_area("💬 Comment", key="chat_cmt",
                 placeholder="Chat QA 특이사항을 입력하세요.")
    st.text_area("📋 02월 목표 및 계획", key="chat_plan",
                 placeholder="목표 및 Action Plan을 입력하세요.")

# ══════════════════════════════════════════════════
# TAB3 : 직무테스트
# ══════════════════════════════════════════════════
with tab3:
    st.markdown("### 📝 직무테스트 결과")

    with st.expander("📋 데이터 입력 방법", expanded=False):
        st.markdown("""
        **필요 컬럼:** `고유번호`, `월`, `팀`, `직무`, `상담사`, `입사일`, `근속개월`, `점수`  
        엑셀에서 전체 데이터 복사(Ctrl+A → Ctrl+C) 후 붙여넣기
        """)

    test_txt = st.text_area(
        "직무테스트 데이터 붙여넣기",
        height=200,
        placeholder="고유번호\t월\t팀\t직무\t상담사\t입사일\t근속개월\t점수\n...",
        key="test_data",
        label_visibility="collapsed",
    )

    st.markdown("**📋 문항별 오답률 (선택)**")
    wrong_txt = st.text_area(
        "문항별 오답률",
        height=100,
        placeholder="문제\t1번\t2번\t3번\t4번\t5번\t6번\t7번\t8번\t9번\t10번\n유형\t취소\t취소\t분쟁조절프로세스\t...\n오답률\t10.5%\t47.4%\t78.9%\t63.2%\t...",
        key="wrong_data",
        label_visibility="collapsed",
    )

    if st.button("📊 직무테스트 장표 생성", key="btn_test", type="primary"):
        test_df = parse_tsv(test_txt)

        if test_df.empty:
            box("데이터를 붙여넣기 해주세요.", "alert")
        else:
            # 컬럼 탐지
            score_c  = next((c for c in test_df.columns if "점수" in c), None)
            month_c  = next((c for c in test_df.columns if "월" in c), None)
            group_c  = next((c for c in test_df.columns if "근속" in c), None)
            agent_c  = next((c for c in test_df.columns if "상담사" in c or "이름" in c), None)
            team_c   = next((c for c in test_df.columns if "팀" in c), None)

            if score_c:
                test_df[score_c] = to_num(test_df[score_c])

            # ── KPI ──
            sec("① KPI 요약")
            curr_scores = to_num(test_df[score_c]).dropna() if score_c else pd.Series()
            curr_mean   = float(curr_scores.mean()) if len(curr_scores) else 0
            pass_cnt    = int((curr_scores >= test_pass).sum())
            fail_cnt    = int((curr_scores < test_pass).sum())
            pass_rate   = pass_cnt / len(curr_scores) * 100 if len(curr_scores) else 0
            total_n     = len(curr_scores)

            c_k = st.columns(5)
            c_k[0].markdown(kpi("평균점수",   curr_mean, sub=f"기준 {test_pass}점"), unsafe_allow_html=True)
            c_k[1].markdown(kpi("응시인원",   total_n,   fmt=",d", sub="명"),        unsafe_allow_html=True)
            c_k[2].markdown(kpi("합격인원",   pass_cnt,  fmt=",d", sub="명"),        unsafe_allow_html=True)
            c_k[3].markdown(kpi("불합격인원", fail_cnt,  fmt=",d", sub="명"),        unsafe_allow_html=True)
            c_k[4].markdown(kpi("합격률",     pass_rate, suffix="%"),               unsafe_allow_html=True)

            if pass_rate < 50:
                box(f"⚠️ 합격률이 <b>{pass_rate:.1f}%</b>로 낮습니다. 집중 보수교육이 필요합니다.", "danger")
            elif pass_rate < 70:
                box(f"📌 합격률 <b>{pass_rate:.1f}%</b>. 취약 인원 보강이 필요합니다.", "alert")
            else:
                box(f"✅ 합격률 <b>{pass_rate:.1f}%</b>.", "good")

            st.markdown("---")

            # ── 차트 ──
            c1t, c2t = st.columns(2)

            with c1t:
                sec("② 월별 평균점수 추이 (근속그룹)")
                if month_c and group_c and score_c:
                    trend = test_df.groupby([month_c, group_c])[score_c].mean().reset_index()
                    trend.columns = ["월", "근속그룹", "평균점수"]
                    trend["평균점수"] = trend["평균점수"].round(1)

                    fig6 = px.bar(
                        trend, x="월", y="평균점수", color="근속그룹",
                        barmode="group", text="평균점수",
                        color_discrete_sequence=COLORS,
                        title="월별 근속그룹별 평균점수",
                    )
                    # 전체 평균 라인
                    overall = test_df.groupby(month_c)[score_c].mean().reset_index()
                    overall.columns = ["월","전체평균"]
                    fig6.add_trace(go.Scatter(
                        x=overall["월"], y=overall["전체평균"].round(1),
                        mode="lines+markers+text",
                        name="전체평균",
                        line=dict(color="#ffa726", width=3),
                        text=overall["전체평균"].round(1),
                        textposition="top center",
                    ))
                    fig6.add_hline(y=test_pass, line_dash="dot", line_color="#ef5350",
                                   annotation_text=f"기준 {test_pass}점")
                    fig6.update_traces(texttemplate="%{text:.1f}",
                                       textposition="outside",
                                       selector=dict(type="bar"))
                    fig6.update_layout(**CHART_BASE, height=380,
                                       yaxis=dict(range=[0,110]),
                                       legend=dict(orientation="h", y=1.08))
                    st.plotly_chart(fig6, use_container_width=True)

            with c2t:
                sec("③ 점수대별 인원 분포")
                if score_c:
                    bins   = [0,40,60,80,90,100,101]
                    labels = ["0~40","40~60","60~80","80","90","100"]
                    test_df["점수구간"] = pd.cut(
                        test_df[score_c], bins=bins, labels=labels,
                        right=False, include_lowest=True
                    )
                    dist = test_df.groupby(["점수구간", group_c] if group_c else ["점수구간"]).size().reset_index(name="인원")
                    if group_c:
                        fig7 = px.bar(dist, x="점수구간", y="인원",
                                      color=group_c, barmode="group",
                                      text="인원",
                                      color_discrete_sequence=COLORS,
                                      title="점수대별 인원 분포")
                    else:
                        fig7 = px.bar(dist, x="점수구간", y="인원",
                                      text="인원", color_discrete_sequence=["#3949ab"],
                                      title="점수대별 인원 분포")
                    fig7.add_vline(x=2.5, line_dash="dot", line_color="#ef5350",
                                   annotation_text=f"기준 {test_pass}점")
                    fig7.update_traces(textposition="outside")
                    fig7.update_layout(**CHART_BASE, height=380,
                                       legend=dict(orientation="h", y=1.08))
                    st.plotly_chart(fig7, use_container_width=True)

            # ── 근속그룹별 현황표 ──
            sec("④ 근속그룹별 현황표")
            if group_c and month_c and score_c:
                # 당월 / 전월 분리
                months = sorted(test_df[month_c].unique())
                curr_m = months[-1] if months else None
                prev_m = months[-2] if len(months) >= 2 else None

                def make_summary(sub_df, label):
                    if sub_df.empty: return pd.DataFrame()
                    s = sub_df.groupby(group_c)[score_c].agg(
                        평균=lambda x: round(safe_mean(x),1),
                        인원=lambda x: int(x.count()),
                    ).reset_index()
                    s.columns = [group_c, f"{label}_평균", f"{label}_인원"]
                    return s

                curr_s = make_summary(test_df[test_df[month_c]==curr_m], "당월") if curr_m else pd.DataFrame()
                prev_s = make_summary(test_df[test_df[month_c]==prev_m], "전월") if prev_m else pd.DataFrame()

                if not curr_s.empty and not prev_s.empty:
                    summary = curr_s.merge(prev_s, on=group_c, how="outer")
                    summary["평균증감"] = (
                        to_num(summary["당월_평균"]) - to_num(summary["전월_평균"])
                    ).round(1).apply(lambda v: f"{v:+.1f}" if not pd.isna(v) else "-")
                    st.dataframe(summary, use_container_width=True, hide_index=True)
                else:
                    st.dataframe(curr_s if not curr_s.empty else test_df,
                                 use_container_width=True, hide_index=True)

            # ── 개인별 점수표 ──
            sec("⑤ 개인별 점수 현황")
            if score_c and agent_c:
                cols_show = [c for c in [month_c, team_c, group_c, agent_c, "직무" if "직무" in test_df.columns else None, score_c] if c and c in test_df.columns]
                indiv = test_df[cols_show].copy()
                if score_c in indiv.columns:
                    indiv["합격여부"] = indiv[score_c].apply(
                        lambda v: "✅ 합격" if pd.notna(v) and v >= test_pass else "❌ 불합격"
                    )
                st.dataframe(
                    indiv.sort_values(score_c, ascending=False),
                    use_container_width=True, hide_index=True, height=400
                )

            # ── 문항별 오답률 ──
            if wrong_txt.strip():
                sec("⑥ 문항별 오답률")
                wrong_df = parse_tsv(wrong_txt)
                if not wrong_df.empty:
                    # 오답률 행 탐지
                    rate_row = wrong_df[wrong_df.iloc[:,0].str.contains("오답률|오답|rate",na=False,case=False)]
                    type_row = wrong_df[wrong_df.iloc[:,0].str.contains("유형|type",na=False,case=False)]

                    if not rate_row.empty:
                        num_cols_w = [c for c in wrong_df.columns if c != wrong_df.columns[0]]
                        rates = rate_row.iloc[0][num_cols_w]
                        rates_num = to_num(rates.astype(str).str.replace("%",""))
                        types = type_row.iloc[0][num_cols_w].values if not type_row.empty else num_cols_w

                        wr_df = pd.DataFrame({
                            "문항": num_cols_w,
                            "유형": types,
                            "오답률": rates_num.values,
                        }).dropna(subset=["오답률"])

                        colors_w = ["#ef5350" if v >= 50 else "#ffa726" if v >= 30 else "#3949ab"
                                    for v in wr_df["오답률"]]
                        fig8 = go.Figure(go.Bar(
                            x=wr_df["문항"], y=wr_df["오답률"],
                            marker_color=colors_w,
                            text=wr_df["오답률"].round(1).astype(str) + "%",
                            textposition="outside",
                            customdata=wr_df["유형"],
                            hovertemplate="%{x}<br>유형: %{customdata}<br>오답률: %{y:.1f}%<extra></extra>",
                        ))
                        fig8.add_hline(y=50, line_dash="dot", line_color="#ef5350",
                                       annotation_text="50% 기준")
                        fig8.update_layout(**CHART_BASE,
                                           title="문항별 오답률",
                                           height=350,
                                           yaxis=dict(range=[0,110], title="오답률(%)"))
                        st.plotly_chart(fig8, use_container_width=True)
                        st.dataframe(wrong_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.text_area("💬 Comment", key="test_cmt",
                 placeholder="직무테스트 특이사항을 입력하세요.")
    st.text_area("📋 02월 목표 및 계획", key="test_plan",
                 placeholder="보수교육 계획 등을 입력하세요.")

# ══════════════════════════════════════════════════
# TAB4 : CSAT
# ══════════════════════════════════════════════════
with tab4:
    st.markdown("### ⭐ CSAT 결과")

    with st.expander("📋 데이터 입력 방법", expanded=False):
        st.markdown("""
        **현황 데이터 컬럼:** `구분`, `전화_12월`, `전화_01월`, `채팅_12월`, `채팅_01월`, `전체_12월`, `전체_01월`, `GAP`  
        **70점 이하 모니터링 컬럼:** `귀책사유`, `전화_12월건`, `전화_12월모니`, `전화_01월건`, `전화_01월모니`,  
        `채팅_12월건`, `채팅_12월모니`, `채팅_01월건`, `채팅_01월모니`, `전체_12월건`, `전체_12월모니`, `전체_01월건`, `전체_01월모니`, `모니터링GAP`
        """)

    st.markdown("**📊 CSAT 현황 데이터**")
    csat_main_txt = st.text_area(
        "CSAT 현황",
        height=140,
        placeholder="구분\t전화_12월\t전화_01월\t채팅_12월\t채팅_01월\t전체_12월\t전체_01월\tGAP\n친절점수\t95.8\t92.8\t92.1\t90.2\t95.2\t92.5\t-2.9%▼\n만족점수\t93.8\t91.8\t86.3\t88.2\t92.6\t91.4\t-1.3%▼\n전체\t94.8\t92.3\t89.2\t89.2\t93.9\t92.0\t-2.1%▼",
        key="csat_main",
        label_visibility="collapsed",
    )

    st.markdown("**🔢 발송/회신 건수 데이터 (선택)**")
    csat_count_txt = st.text_area(
        "발송회신 건수",
        height=100,
        placeholder="구분\t전화_12월\t전화_01월\t채팅_12월\t채팅_01월\t전체_12월\t전체_01월\tGAP\n발송\t3583\t2877\t520\t285\t4103\t3162\t-22.9%▼\n회신\t411\t368\t76\t51\t487\t419\t-14.0%▼\n회신율\t11.5%\t12.8%\t14.6%\t17.9%\t11.9%\t13.3%\t+1.3%▲",
        key="csat_count",
        label_visibility="collapsed",
    )

    st.markdown("**🔴 70점 이하 모니터링 데이터**")
    csat_low_txt = st.text_area(
        "70점 이하 모니터링",
        height=160,
        placeholder="귀책사유\t전화_12월건\t전화_12월모니\t전화_01월건\t전화_01월모니\t채팅_12월건\t채팅_12월모니\t채팅_01월건\t채팅_01월모니\t전체_12월건\t전체_12월모니\t전체_01월건\t전체_01월모니\t모니터링GAP\nIBR\t9\t92.2\t18\t88.9\t8\t88.8\t2\t77.5\t17\t90.6\t20\t87.8\t3.1%▼\n고객\t7\t95.7\t11\t98.2\t1\t90.0\t2\t95.0\t8\t95.0\t13\t97.7\t2.8%▲\n상담사\t10\t79.0\t12\t80.4\t4\t85.0\t4\t90.0\t14\t80.7\t16\t82.8\t2.6%▲\n합계/평균\t26\t88.1\t41\t88.9\t13\t87.7\t8\t88.1\t39\t87.9\t49\t88.8\t1.0%▲",
        key="csat_low",
        label_visibility="collapsed",
    )

    if st.button("📊 CSAT 장표 생성", key="btn_csat", type="primary"):
        main_df  = parse_tsv(csat_main_txt)
        count_df = parse_tsv(csat_count_txt)
        low_df   = parse_tsv(csat_low_txt)

        if main_df.empty and low_df.empty:
            box("데이터를 붙여넣기 해주세요.", "alert")
        else:
            # ── KPI ──
            sec("① CSAT KPI 요약")
            전체행 = main_df[main_df.iloc[:,0].str.contains("전체|Total",na=False,case=False)] if not main_df.empty else pd.DataFrame()

            # 당월/전월 컬럼 탐지
            curr_cols = [c for c in main_df.columns if report_ym[-3:] in c or "01월" in c] if not main_df.empty else []
            prev_cols = [c for c in main_df.columns if prev_ym[-3:]  in c or "12월" in c] if not main_df.empty else []

            curr_val = safe_mean(to_num(전체행[curr_cols[0]])) if (not 전체행.empty and curr_cols) else 0
            prev_val = safe_mean(to_num(전체행[prev_cols[0]])) if (not 전체행.empty and prev_cols) else 0
            delta_c  = curr_val - prev_val if prev_val else None

            # 회신율
            회신율_row = count_df[count_df.iloc[:,0].str.contains("회신율",na=False)] if not count_df.empty else pd.DataFrame()
            curr_rate  = 0.0
            if not 회신율_row.empty and curr_cols:
                rv = 회신율_row.iloc[0][curr_cols[0]] if curr_cols[0] in 회신율_row.columns else "0"
                curr_rate = float(str(rv).replace("%","").replace("▲","").replace("▼","")) if rv else 0

            # 70점 이하 건수
            합계행 = low_df[low_df.iloc[:,0].str.contains("합계|평균|Total",na=False,case=False)] if not low_df.empty else pd.DataFrame()
            low_curr_cols = [c for c in low_df.columns if "01월건" in c or "전체_01월건" in c] if not low_df.empty else []
            low_cnt = int(safe_mean(to_num(합계행[low_curr_cols[0]]))) if (not 합계행.empty and low_curr_cols) else 0

            c_k2 = st.columns(5)
            c_k2[0].markdown(kpi("전체 CSAT",   curr_val,   sub=f"목표 {csat_target}점", delta=delta_c), unsafe_allow_html=True)
            c_k2[1].markdown(kpi("전월 CSAT",   prev_val,   fmt=".1f"), unsafe_allow_html=True)
            c_k2[2].markdown(kpi("GAP",          (curr_val-prev_val) if prev_val else 0, fmt="+.1f"), unsafe_allow_html=True)
            c_k2[3].markdown(kpi("회신율",        curr_rate,  suffix="%"), unsafe_allow_html=True)
            c_k2[4].markdown(kpi("70점이하 건수", low_cnt,    fmt=",d", sub="건"), unsafe_allow_html=True)

            if curr_val < csat_target:
                box(f"⚠️ CSAT <b>{curr_val:.1f}점</b>으로 목표 <b>{csat_target}점</b> 미달 (GAP {curr_val-csat_target:+.1f}점)", "danger")
            else:
                box(f"✅ CSAT <b>{curr_val:.1f}점</b>으로 목표 달성!", "good")

            st.markdown("---")

            # ── CSAT 현황표 + 차트 ──
            c1s, c2s = st.columns([2, 3])

            with c1s:
                sec("② CSAT 현황표")
                if not main_df.empty:
                    st.dataframe(main_df, use_container_width=True, hide_index=True)
                if not count_df.empty:
                    st.markdown("**발송/회신 현황**")
                    st.dataframe(count_df, use_container_width=True, hide_index=True)

            with c2s:
                sec("③ 전화/채팅/전체 비교 차트")
                if not main_df.empty and curr_cols and prev_cols:
                    구분_c = main_df.columns[0]
                    score_rows = main_df[~main_df[구분_c].str.contains("합계|Total",na=False,case=False)].copy()

                    # long format 변환
                    plot_cols_s = [c for c in main_df.columns if "전화" in c or "채팅" in c or "전체" in c]
                    if plot_cols_s:
                        melt_s = score_rows[[구분_c]+plot_cols_s].melt(
                            id_vars=구분_c, var_name="채널_월", value_name="점수")
                        melt_s["점수"] = to_num(melt_s["점수"].astype(str).str.replace("[▲▼%]","",regex=True))
                        melt_s = melt_s.dropna(subset=["점수"])

                        fig9 = px.bar(
                            melt_s, x="채널_월", y="점수",
                            color=구분_c, barmode="group",
                            text="점수",
                            color_discrete_sequence=COLORS,
                            title="채널별 CSAT 비교",
                        )
                        fig9.add_hline(y=csat_target, line_dash="dot", line_color="#ef5350",
                                       annotation_text=f"목표 {csat_target}점")
                        fig9.update_traces(texttemplate="%{text:.1f}", textposition="outside")
                        fig9.update_layout(**CHART_BASE, height=380,
                                           yaxis=dict(range=[80,100]),
                                           xaxis=dict(tickangle=-30),
                                           legend=dict(orientation="h", y=1.08))
                        st.plotly_chart(fig9, use_container_width=True)

            # ── 70점 이하 모니터링 ──
            sec("④ 70점 이하 모니터링 결과")
            if not low_df.empty:
                st.dataframe(low_df, use_container_width=True, hide_index=True)

                st.markdown("---")
                c3s, c4s = st.columns(2)

                with c3s:
                    # 귀책별 파이차트 (당월 건수)
                    귀책_c = low_df.columns[0]
                    건수_cols = [c for c in low_df.columns if "01월건" in c or "건" in c]
                    data_rows_low = low_df[~low_df[귀책_c].str.contains("합계|평균|Total",na=False,case=False)]

                    if 건수_cols and not data_rows_low.empty:
                        pie_data = data_rows_low[[귀책_c, 건수_cols[0]]].copy()
                        pie_data.columns = ["귀책사유","건수"]
                        pie_data["건수"] = to_num(pie_data["건수"])
                        pie_data = pie_data.dropna()

                        fig10 = px.pie(
                            pie_data, values="건수", names="귀책사유",
                            title=f"{report_ym} 귀책사유별 건수",
                            color="귀책사유",
                            color_discrete_sequence=COLORS,
                            hole=0.4,
                        )
                        fig10.update_traces(textinfo="percent+label+value")
                        fig10.update_layout(**CHART_BASE, height=350)
                        st.plotly_chart(fig10, use_container_width=True)

                with c4s:
                    # 귀책별 모니터링 점수 비교 (전월/당월)
                    prev_moni_cols = [c for c in low_df.columns if "12월모니" in c or "전월모니" in c]
                    curr_moni_cols = [c for c in low_df.columns if "01월모니" in c or "당월모니" in c]

                    if prev_moni_cols and curr_moni_cols and not data_rows_low.empty:
                        귀책_list = data_rows_low[귀책_c].tolist()

                        # 전체 컬럼 합산
                        prev_moni_vals = data_rows_low[prev_moni_cols].apply(
                            lambda r: to_num(r).mean(), axis=1).round(1).tolist()
                        curr_moni_vals = data_rows_low[curr_moni_cols].apply(
                            lambda r: to_num(r).mean(), axis=1).round(1).tolist()

                        fig11 = go.Figure()
                        fig11.add_trace(go.Bar(
                            name=prev_ym, x=귀책_list, y=prev_moni_vals,
                            marker_color="#42a5f5", text=prev_moni_vals, textposition="outside",
                        ))
                        fig11.add_trace(go.Bar(
                            name=report_ym, x=귀책_list, y=curr_moni_vals,
                            marker_color="#1a237e", text=curr_moni_vals, textposition="outside",
                        ))
                        fig11.add_hline(y=csat_monitor, line_dash="dot", line_color="#ef5350",
                                        annotation_text=f"모니터링 기준 {csat_monitor}점")
                        fig11.update_layout(
                            **CHART_BASE,
                            title="귀책별 모니터링 점수 비교",
                            barmode="group", height=350,
                            yaxis=dict(range=[70,105]),
                            legend=dict(orientation="h", y=1.08),
                        )
                        st.plotly_chart(fig11, use_container_width=True)

    st.markdown("---")
    st.text_area("💬 Comment", key="csat_cmt",
                 placeholder="CSAT 특이사항을 입력하세요.")
    st.text_area("📋 02월 목표 및 계획", key="csat_plan",
                 placeholder="관리 계획을 입력하세요.")

# ─────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────
st.markdown("---")
st.caption(f"📊 교육품질 통합 대시보드 · {report_ym} 기준 · Powered by Streamlit & Plotly")
