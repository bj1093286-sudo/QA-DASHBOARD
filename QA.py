# ============================================================
# CSAT 품질 대시보드 v1.0
# save as: csat_app.py
# run:     streamlit run csat_app.py
# install: pip install streamlit plotly pandas numpy openpyxl
# ============================================================

import io, re
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
    page_title="CSAT 품질 대시보드",
    page_icon="⭐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# 1. GLOBAL CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
.stApp { background: #f4f6fb; }

.kpi-card {
    background: #fff;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 2px 10px rgba(30,40,90,.10);
    text-align: center;
    margin-bottom: 8px;
}
.kpi-label  { font-size:.75rem; color:#6b7a99; font-weight:700; letter-spacing:.5px; text-transform:uppercase; }
.kpi-value  { font-size:2rem; font-weight:800; color:#1a237e; line-height:1.2; margin:4px 0; }
.kpi-sub    { font-size:.78rem; color:#888; }
.kpi-up     { color:#1b8a4c; font-size:.82rem; }
.kpi-down   { color:#c0392b; font-size:.82rem; }
.kpi-eq     { color:#888;    font-size:.82rem; }

.sec-header {
    background: linear-gradient(90deg,#1a237e,#3949ab);
    color:#fff; border-radius:8px;
    padding:9px 16px; font-size:1rem;
    font-weight:700; margin:20px 0 10px;
}

.alert-box {
    background:#fff3cd; border-left:4px solid #ffc107;
    border-radius:6px; padding:10px 14px;
    margin:8px 0; font-size:.88rem;
}
.danger-box {
    background:#fdecea; border-left:4px solid #ef5350;
    border-radius:6px; padding:10px 14px;
    margin:8px 0; font-size:.88rem;
}
.good-box {
    background:#e8f5e9; border-left:4px solid #43a047;
    border-radius:6px; padding:10px 14px;
    margin:8px 0; font-size:.88rem;
}

div[data-baseweb="tab"] {
    font-weight:600; border-radius:8px 8px 0 0; padding:8px 18px;
}

table { width:100%; border-collapse:collapse; font-size:.85rem; }
th { background:#1a237e; color:#fff; padding:8px 10px; text-align:center; }
td { padding:7px 10px; border-bottom:1px solid #eee; text-align:center; }
tr:hover td { background:#f0f4ff; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 2. HELPERS
# ─────────────────────────────────────────
COLORS = ["#1a237e","#3949ab","#42a5f5","#26c6da",
          "#66bb6a","#ffa726","#ef5350","#ab47bc","#ec407a","#26a69a"]

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

def safe_mean(s):
    v = to_num(s).dropna()
    return float(v.mean()) if len(v) else 0.0

def clean_col(c):
    return re.sub(r"\s+", "", str(c)).strip()

def kpi(label, value, sub=None, delta=None, fmt=".1f", suffix=""):
    v = f"{value:{fmt}}{suffix}" if isinstance(value,(int,float)) else str(value)
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    if delta is None:   d=""
    elif delta > 0:     d=f'<div class="kpi-up">▲ {delta:+.1f}</div>'
    elif delta < 0:     d=f'<div class="kpi-down">▼ {delta:+.1f}</div>'
    else:               d='<div class="kpi-eq">― 변동없음</div>'
    return f"""<div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{v}</div>
      {sub_html}{d}
    </div>"""

def sec(txt):
    st.markdown(f'<div class="sec-header">{txt}</div>', unsafe_allow_html=True)

def alert(txt, kind="alert"):
    css = {"alert":"alert-box","danger":"danger-box","good":"good-box"}.get(kind,"alert-box")
    st.markdown(f'<div class="{css}">{txt}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# 3. DATA PARSER
# ─────────────────────────────────────────
def parse_file(file) -> pd.DataFrame:
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            for enc in ["utf-8-sig","euc-kr","cp949","utf-8"]:
                try:
                    df = pd.read_csv(file, encoding=enc)
                    file.seek(0)
                    return df
                except Exception:
                    file.seek(0)
        else:
            # 엑셀: 헤더가 병합/다중일 수 있으므로 raw로 읽고 탐지
            raw = pd.read_excel(file, header=None, dtype=str)
            # "회신" 또는 "상담사" 포함 행을 헤더로
            for i, row in raw.iterrows():
                vals = [str(v) for v in row]
                if any("회신" in v or "상담사" in v or "채널" in v for v in vals):
                    # 헤더 2행 병합 처리 (다음 행도 서브헤더일 수 있음)
                    header_row = raw.iloc[i].fillna("").astype(str)
                    # 다음 행이 서브헤더인지 확인
                    next_row = raw.iloc[i+1].fillna("").astype(str) if i+1 < len(raw) else None
                    if next_row is not None:
                        # 서브헤더 병합
                        merged = []
                        for h, s in zip(header_row, next_row):
                            h = h.strip()
                            s = s.strip()
                            if s and s not in h:
                                merged.append(f"{h}_{s}" if h else s)
                            else:
                                merged.append(h)
                        df = raw.iloc[i+2:].copy()
                        df.columns = merged
                    else:
                        df = raw.iloc[i+1:].copy()
                        df.columns = header_row.tolist()
                    df = df.reset_index(drop=True)
                    return df
            # 헤더 탐지 실패 시 그냥 반환
            df = pd.read_excel(file, dtype=str)
            return df
    except Exception as e:
        st.error(f"파일 파싱 오류: {e}")
        return pd.DataFrame()

def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼명 정제 + 핵심 컬럼 매핑"""
    # 컬럼명 공백 제거
    df.columns = [clean_col(c) for c in df.columns]
    df = df.dropna(how="all").copy()

    # ── 컬럼 별칭 매핑 ──────────────────────────────
    alias = {
        "회신월":    ["회신월","월"],
        "회신주차":  ["회신주차","주차","week","wk"],
        "회신일자":  ["회신일자","일자","날짜","date"],
        "상담사":    ["상담사"],
        "채널":      ["채널","channel"],
        "키워드":    ["키워드","keyword"],
        "긍부정":    ["긍정/부정","긍부정","sentiment"],
        "유형":      ["유형","type"],
        "총합":      ["총합","totalscore","total"],
        "Q1":        ["Q1"],
        "Q2":        ["Q2"],
        "Q3":        ["Q3"],
        "친절점수":  ["친절점수","친절"],
        "만족점수":  ["만족점수","만족"],
        "최종점수":  ["최종점수","finalScore","이행점수"],
        "만족율":    ["만족율(건)","만족율","satisfactionrate"],
        "상담KEY":   ["상담KEY","상담key","KEY","key"],
        "문의유형":  ["문의유형","inquirytype"],
        "귀책분류":  ["귀책분류","responsibility"],
        "문의불만사유":["문의불만사유","불만사유","reason"],
        "상세분석":  ["상세분석","analysis","분석"],
        "상담사근속":["상담사근속","상담사_근속","근속","근속그룹","tenure"],
        "피드백여부":["피드백여부","피드백_여부","feedback"],
        "피드백결과":["피드백결과","피드백_결과","feedbackresult"],
    }

    col_map = {}
    for std, candidates in alias.items():
        for c in df.columns:
            if any(clean_col(cand).lower() in c.lower() or c.lower() in clean_col(cand).lower()
                   for cand in candidates):
                col_map[c] = std
                break

    df = df.rename(columns=col_map)

    # ── 날짜/주차 처리 ──────────────────────────────
    if "회신일자" in df.columns:
        df["회신일자"] = pd.to_datetime(df["회신일자"], errors="coerce")
    if "회신주차" not in df.columns and "회신일자" in df.columns:
        df["회신주차"] = df["회신일자"].apply(
            lambda x: f"{x.month}월{((x.day-1)//7)+1}주" if pd.notna(x) else "미상"
        )

    # ── 숫자 변환 ──────────────────────────────────
    for col in ["총합","Q1","Q2","Q3","친절점수","만족점수","최종점수"]:
        if col in df.columns:
            df[col] = to_num(df[col])

    # ── 만족율 처리 ─────────────────────────────────
    if "만족율" in df.columns:
        def parse_pct(v):
            v = str(v).strip().replace("%","")
            try: return float(v)*100 if float(v)<=1 else float(v)
            except: return np.nan
        df["만족율_num"] = df["만족율"].apply(parse_pct)
    else:
        df["만족율_num"] = np.nan

    # ── 70점 미만 플래그 ────────────────────────────
    score_col = "최종점수" if "최종점수" in df.columns else (
        "만족점수" if "만족점수" in df.columns else None)
    if score_col:
        df["_score"] = df[score_col]
        df["_below70"] = df["_score"] < 70
    else:
        df["_score"] = np.nan
        df["_below70"] = False

    # ── 채널 정규화 ────────────────────────────────
    if "채널" in df.columns:
        df["채널_구분"] = df["채널"].apply(
            lambda x: "채팅" if "채팅" in str(x) or "chat" in str(x).lower()
            else ("전화" if "전화" in str(x) or "call" in str(x).lower() or "IN" in str(x) else str(x))
        )
    else:
        df["채널_구분"] = "전체"

    return df

# ─────────────────────────────────────────
# 4. CHART HELPERS
# ─────────────────────────────────────────
CHART_LAYOUT = dict(
    plot_bgcolor="#fff", paper_bgcolor="#fff",
    font=dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif", size=12),
    margin=dict(t=50, b=40, l=40, r=20),
)

def fig_weekly_line(df, score_col, group_col, title, target=70):
    """주차별 그룹 라인차트"""
    if score_col not in df.columns or group_col not in df.columns:
        return None
    wk_col = "회신주차" if "회신주차" in df.columns else None
    if not wk_col:
        return None
    g = df.groupby([wk_col, group_col])[score_col].mean().reset_index()
    g.columns = ["주차", group_col, "평균점수"]

    # 주차 정렬
    def wk_sort(w):
        try:
            m = int(re.search(r"(\d+)월", str(w)).group(1))
            n = int(re.search(r"(\d+)주", str(w)).group(1))
            return m*10+n
        except: return 0
    wk_order = sorted(g["주차"].unique(), key=wk_sort)
    g["주차"] = pd.Categorical(g["주차"], categories=wk_order, ordered=True)
    g = g.sort_values("주차")

    fig = px.line(g, x="주차", y="평균점수", color=group_col,
                  markers=True, title=title,
                  color_discrete_sequence=COLORS)
    fig.add_hline(y=target, line_dash="dot", line_color="#ef5350",
                  annotation_text=f"기준 {target}점",
                  annotation_position="bottom right")
    fig.update_layout(**CHART_LAYOUT, height=360,
                      yaxis=dict(range=[0,105]),
                      legend=dict(orientation="h", y=1.1))
    return fig

def fig_bar_agent(df, score_col, title, threshold=70, orient="h"):
    """상담사별 평균점수 바"""
    if score_col not in df.columns or "상담사" not in df.columns:
        return None
    ag = df.groupby("상담사")[score_col].agg(["mean","count"]).reset_index()
    ag.columns = ["상담사","평균점수","건수"]
    ag["평균점수"] = ag["평균점수"].round(1)
    ag = ag.sort_values("평균점수", ascending=(orient=="h"))

    colors = ["#ef5350" if v < threshold else "#3949ab" for v in ag["평균점수"]]

    if orient == "h":
        fig = go.Figure(go.Bar(
            x=ag["평균점수"], y=ag["상담사"],
            orientation="h", marker_color=colors,
            text=ag["평균점수"], textposition="outside",
            customdata=ag["건수"],
            hovertemplate="%{y}<br>평균: %{x:.1f}점<br>건수: %{customdata}건<extra></extra>",
        ))
        fig.add_vline(x=threshold, line_dash="dot", line_color="#ef5350",
                      annotation_text=f"기준 {threshold}점")
        fig.update_layout(**CHART_LAYOUT,
                          title=title,
                          height=max(300, len(ag)*36),
                          xaxis=dict(range=[0,110]))
    else:
        fig = go.Figure(go.Bar(
            x=ag["상담사"], y=ag["평균점수"],
            marker_color=colors,
            text=ag["평균점수"], textposition="outside",
            customdata=ag["건수"],
            hovertemplate="%{x}<br>평균: %{y:.1f}점<br>건수: %{customdata}건<extra></extra>",
        ))
        fig.add_hline(y=threshold, line_dash="dot", line_color="#ef5350",
                      annotation_text=f"기준 {threshold}점")
        fig.update_layout(**CHART_LAYOUT,
                          title=title,
                          height=380,
                          yaxis=dict(range=[0,110]))
    return fig

def fig_channel_bar(df, score_col, title):
    """채널별 평균점수"""
    if score_col not in df.columns:
        return None
    ch = df.groupby("채널_구분")[score_col].agg(["mean","count"]).reset_index()
    ch.columns = ["채널","평균점수","건수"]
    ch["평균점수"] = ch["평균점수"].round(1)

    fig = px.bar(ch, x="채널", y="평균점수",
                 color="채널", text="평균점수",
                 color_discrete_sequence=COLORS,
                 title=title)
    fig.update_traces(texttemplate="%{text:.1f}점", textposition="outside")
    fig.update_layout(**CHART_LAYOUT, height=320,
                      yaxis=dict(range=[0,110]), showlegend=False)
    return fig

def fig_sentiment_pie(df, title="긍정/부정 비율"):
    """긍부정 파이차트"""
    if "긍부정" not in df.columns:
        return None
    vc = df["긍부정"].value_counts().reset_index()
    vc.columns = ["구분","건수"]
    color_map = {"긍정":"#43a047","부정":"#ef5350","중립":"#ffa726"}
    fig = px.pie(vc, values="건수", names="구분",
                 title=title,
                 color="구분",
                 color_discrete_map=color_map,
                 hole=0.4)
    fig.update_traces(textinfo="percent+label",
                      textfont_size=13)
    fig.update_layout(**CHART_LAYOUT, height=320)
    return fig

def fig_reason_bar(df, col, title, top_n=10):
    """불만사유 / 유형 빈도"""
    if col not in df.columns:
        return None
    vc = df[col].value_counts().head(top_n).reset_index()
    vc.columns = ["항목","건수"]
    fig = px.bar(vc, x="건수", y="항목",
                 orientation="h",
                 color="건수",
                 color_continuous_scale="Blues",
                 title=title,
                 text="건수")
    fig.update_traces(textposition="outside")
    fig.update_layout(**CHART_LAYOUT,
                      height=max(280, len(vc)*34),
                      showlegend=False,
                      coloraxis_showscale=False)
    return fig

def fig_score_hist(df, score_col, title, threshold=70):
    """점수 분포 히스토그램"""
    if score_col not in df.columns:
        return None
    scores = to_num(df[score_col]).dropna()
    fig = go.Figure()
    fig.add_trace(go.Histogram(
        x=scores, nbinsx=20,
        marker_color="#3949ab",
        name="점수분포",
        opacity=0.85,
    ))
    fig.add_vline(x=threshold, line_dash="dot", line_color="#ef5350",
                  annotation_text=f"기준 {threshold}점")
    fig.update_layout(**CHART_LAYOUT,
                      title=title, height=300,
                      xaxis_title="점수", yaxis_title="건수")
    return fig

def fig_area_heatmap(df, area_cols, group_col, title):
    """영역별 점수 히트맵"""
    valid = [c for c in area_cols if c in df.columns]
    if not valid or group_col not in df.columns:
        return None
    heat = df.groupby(group_col)[valid].mean().round(1)
    fig = go.Figure(go.Heatmap(
        z=heat.values,
        x=heat.columns.tolist(),
        y=heat.index.tolist(),
        colorscale="Blues",
        text=heat.values.round(1),
        texttemplate="%{text}",
        textfont={"size":11},
        hoverongaps=False,
    ))
    fig.update_layout(**CHART_LAYOUT, title=title, height=350)
    return fig

def fig_weekly_satisfaction(df, title):
    """주차별 만족율(%) 라인차트"""
    if "만족율_num" not in df.columns or "회신주차" not in df.columns:
        return None
    g = df.groupby("회신주차")["만족율_num"].mean().reset_index()

    def wk_sort(w):
        try:
            m = int(re.search(r"(\d+)월", str(w)).group(1))
            n = int(re.search(r"(\d+)주", str(w)).group(1))
            return m*10+n
        except: return 0
    g = g.sort_values("회신주차", key=lambda s: s.map(wk_sort))

    fig = px.line(g, x="회신주차", y="만족율_num",
                  markers=True, title=title,
                  color_discrete_sequence=["#1a237e"])
    fig.add_hline(y=50, line_dash="dot", line_color="#ffa726",
                  annotation_text="50% 기준")
    fig.update_layout(**CHART_LAYOUT, height=300,
                      yaxis_title="만족율(%)",
                      yaxis=dict(range=[0,110]))
    return fig

# ─────────────────────────────────────────
# 5. SIDEBAR
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⭐ CSAT 대시보드")
    st.markdown("---")
    report_month   = st.text_input("📅 보고 기준 월", value="2025년 12월")
    score_threshold= st.number_input("🔴 모니터링 기준점수", value=70.0, step=1.0,
                                      help="이 점수 미만은 모니터링 대상으로 표시됩니다.")
    target_sat     = st.number_input("🎯 만족율 목표(%)", value=70.0, step=1.0)
    st.markdown("---")
    st.caption("📌 지원형식: Excel(.xlsx/.xls), CSV(.csv)")
    st.caption("컬럼: 회신월/주차/일자, 상담사, 채널, 최종점수, 귀책분류 등")

# ─────────────────────────────────────────
# 6. HEADER
# ─────────────────────────────────────────
st.markdown(f"""
<div style="background:linear-gradient(90deg,#1a237e,#3949ab);
     color:#fff;border-radius:12px;padding:22px 28px;margin-bottom:20px;">
  <h2 style="margin:0;font-size:1.6rem;">⭐ CSAT 고객만족도 품질 대시보드</h2>
  <p style="margin:4px 0 0;opacity:.85;font-size:.93rem;">
    {report_month} 보고 · 모니터링 기준 {score_threshold:.0f}점 미만 · 만족율 목표 {target_sat:.0f}%
  </p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# 7. FILE UPLOAD
# ─────────────────────────────────────────
uploaded = st.file_uploader(
    "📂 CSAT 원본 데이터 파일 업로드 (Excel / CSV)",
    type=["xlsx","xls","csv"],
    help="회신월, 회신주차, 상담사, 채널, 최종점수, 귀책분류 등 컬럼 포함 파일",
)

if uploaded is None:
    st.info("👆 파일을 업로드하면 CSAT 분석 대시보드가 자동으로 생성됩니다.")

    # 샘플 데이터 형식 안내
    with st.expander("📋 데이터 형식 예시 보기"):
        sample = pd.DataFrame({
            "회신월":["12월","12월","12월"],
            "회신주차":["49WK","49WK","50WK"],
            "회신일자":["2025-12-01","2025-12-02","2025-12-08"],
            "상담사":["홍길동","김철수","이영희"],
            "채널":["전화 IN","채팅","전화 IN"],
            "긍정/부정":["부정","긍정","부정"],
            "유형":["상담사_소극적","상담사_능숙","IBR_시스템"],
            "최종점수":[40,100,75],
            "친절점수":[40,100,60],
            "만족점수":[20,20,60],
            "만족율(건)":["0%","100%","50%"],
            "귀책분류":["IBR","상담사","IBR"],
            "이행점수":[100,100,85],
            "상담사 근속":["기존1(1년이내)","신입4(180일이내)","기존1(1년이내)"],
            "피드백여부":["X","O","O"],
        })
        st.dataframe(sample, use_container_width=True, hide_index=True)
    st.stop()

# ─────────────────────────────────────────
# 8. DATA LOAD & NORMALIZE
# ─────────────────────────────────────────
with st.spinner("📊 데이터 분석 중..."):
    raw = parse_file(uploaded)
    if raw.empty:
        st.error("❌ 데이터를 읽지 못했습니다. 파일 형식을 확인해주세요.")
        st.stop()
    df = normalize(raw)

st.success(f"✅ {len(df):,}건 로드 완료  |  컬럼 수: {len(df.columns)}개  |  상담사: {df['상담사'].nunique() if '상담사' in df.columns else '?'}명")

# 영역 컬럼 탐지 (정확성/숙련도/친절도/약속이행 세부항목)
AREA_COLS = {
    "정확성(30)": [c for c in df.columns if "정확한안내" in c or "프로세스" in c or "전산처리" in c],
    "숙련도(20)": [c for c in df.columns if "맞춤설명" in c or "문의파악" in c or "숙련도" in c],
    "친절도(30)": [c for c in df.columns if "감정연출" in c or "양해" in c or "경청" in c or "호응" in c or "언어표현" in c],
    "약속이행(20)":[c for c in df.columns if "약속" in c or "안내누락" in c],
}

# ─────────────────────────────────────────
# 9. TABS
# ─────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 장표1 · 전체 현황",
    "👤 장표2 · 상담사별 분석",
    "🔴 장표3 · 70점 미만 모니터링",
    "🔍 장표4 · 귀책 & 불만 분석",
])

# ══════════════════════════════════════════
# TAB 1 : 전체 현황
# ══════════════════════════════════════════
with tab1:
    sec("📊 CSAT 전체 현황 요약")

    score_col = "최종점수" if "최종점수" in df.columns else "_score"
    scores = to_num(df[score_col]).dropna()

    total_cnt   = len(df)
    avg_score   = float(scores.mean()) if len(scores) else 0
    below70_cnt = int((scores < score_threshold).sum())
    below70_pct = below70_cnt / total_cnt * 100 if total_cnt else 0
    pos_cnt     = (df["긍부정"] == "긍정").sum() if "긍부정" in df.columns else 0
    neg_cnt     = (df["긍부정"] == "부정").sum() if "긍부정" in df.columns else 0
    sat_avg     = float(df["만족율_num"].mean()) if "만족율_num" in df.columns else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.markdown(kpi("전체 평균점수",    avg_score,  sub=f"기준 {score_threshold:.0f}점"), unsafe_allow_html=True)
    c2.markdown(kpi("총 응대 건수",     total_cnt,  fmt=",d", sub="건"), unsafe_allow_html=True)
    c3.markdown(kpi(f"{score_threshold:.0f}점 미만",  below70_cnt, fmt=",d", sub=f"({below70_pct:.1f}%)"), unsafe_allow_html=True)
    c4.markdown(kpi("만족율 평균",      sat_avg,    suffix="%", sub=f"목표 {target_sat:.0f}%"), unsafe_allow_html=True)
    c5.markdown(kpi("긍정 / 부정",      f"{pos_cnt} / {neg_cnt}", sub="건"), unsafe_allow_html=True)

    # 경고 배너
    if below70_pct >= 30:
        alert(f"⚠️ {score_threshold:.0f}점 미만 비율이 <b>{below70_pct:.1f}%</b>로 높습니다. 즉각 개선 조치가 필요합니다.", "danger")
    elif below70_pct >= 15:
        alert(f"📌 {score_threshold:.0f}점 미만 비율이 <b>{below70_pct:.1f}%</b>입니다. 모니터링이 필요합니다.", "alert")

    st.markdown("---")
    ca, cb = st.columns(2)

    with ca:
        sec("① 주차별 평균 최종점수 추이")
        if "회신주차" in df.columns:
            f = fig_weekly_line(df, score_col, "채널_구분",
                                 "주차별 채널별 평균점수", score_threshold)
            if f: st.plotly_chart(f, use_container_width=True)
        sec("③ 채널별 평균점수")
        f3 = fig_channel_bar(df, score_col, "채널별 평균점수")
        if f3: st.plotly_chart(f3, use_container_width=True)

    with cb:
        sec("② 긍정 / 부정 비율")
        f2 = fig_sentiment_pie(df, "긍정/부정 비율")
        if f2: st.plotly_chart(f2, use_container_width=True)

        sec("④ 주차별 만족율(%)")
        f4 = fig_weekly_satisfaction(df, "주차별 만족율 추이")
        if f4: st.plotly_chart(f4, use_container_width=True)

    sec("⑤ 주차별 상세 현황표")
    if "회신주차" in df.columns:
        wk_tbl = df.groupby("회신주차").agg(
            총건수   =(score_col, "count"),
            평균점수 =(score_col, lambda x: round(safe_mean(x),1)),
            미만건수 =(score_col, lambda x: int((to_num(x)<score_threshold).sum())),
            만족율평균=("만족율_num", lambda x: f"{safe_mean(x):.1f}%"),
            긍정건수 =("긍부정",      lambda x: (x=="긍정").sum() if "긍부정" in df.columns else 0),
            부정건수 =("긍부정",      lambda x: (x=="부정").sum() if "긍부정" in df.columns else 0),
        ).reset_index()
        wk_tbl.columns = ["주차","총건수","평균점수",
                           f"{score_threshold:.0f}점미만","만족율평균","긍정","부정"]
        st.dataframe(wk_tbl, use_container_width=True, hide_index=True)

    st.text_area("💬 Comment", key="t1_cmt",
                 placeholder="전체 CSAT 현황 특이사항을 입력하세요.")
    st.text_area("📋 Action Plan", key="t1_plan",
                 placeholder="전반적인 개선 계획을 입력하세요.")

# ══════════════════════════════════════════
# TAB 2 : 상담사별 분석
# ══════════════════════════════════════════
with tab2:
    sec("👤 상담사별 CSAT 분석")

    if "상담사" not in df.columns:
        st.warning("상담사 컬럼이 없습니다.")
    else:
        score_col = "최종점수" if "최종점수" in df.columns else "_score"

        ca2, cb2 = st.columns(2)
        with ca2:
            sec("① 상담사별 평균 최종점수")
            f = fig_bar_agent(df, score_col, "상담사별 평균 최종점수",
                               threshold=score_threshold, orient="h")
            if f: st.plotly_chart(f, use_container_width=True)

        with cb2:
            sec("② 주차별 상담사 점수 추이")
            f2 = fig_weekly_line(df, score_col, "상담사",
                                  "주차별 상담사 점수 추이", score_threshold)
            if f2: st.plotly_chart(f2, use_container_width=True)

        # ③ 영역별 히트맵
        all_area = [c for cols in AREA_COLS.values() for c in cols]
        if all_area:
            sec("③ 영역별 점수 히트맵 (상담사)")
            f3 = fig_area_heatmap(df, all_area, "상담사", "영역별 점수 히트맵")
            if f3: st.plotly_chart(f3, use_container_width=True)

        # ④ 상담사별 상세표
        sec("④ 상담사별 상세 현황표")
        agg_dict = {
            "총건수":      (score_col, "count"),
            "평균최종점수": (score_col, lambda x: round(safe_mean(x),1)),
            "평균친절점수": ("친절점수", lambda x: round(safe_mean(x),1)) if "친절점수" in df.columns else (score_col, "count"),
            "평균만족점수": ("만족점수", lambda x: round(safe_mean(x),1)) if "만족점수" in df.columns else (score_col, "count"),
            f"{score_threshold:.0f}점미만": (score_col, lambda x: int((to_num(x)<score_threshold).sum())),
            "피드백완료":   ("피드백여부", lambda x: f"{(x=='O').sum()}건") if "피드백여부" in df.columns else (score_col, "count"),
        }
        # 유효한 것만 필터
        valid_agg = {}
        for k, (col, func) in agg_dict.items():
            if col in df.columns:
                valid_agg[k] = (col, func)

        agent_tbl = df.groupby(["상담사근속","상담사"] if "상담사근속" in df.columns else ["상담사"]).agg(
            **{k: v for k, v in valid_agg.items()}
        ).reset_index()

        # 목표달성 여부 추가
        avg_col = "평균최종점수"
        if avg_col in agent_tbl.columns:
            agent_tbl["목표달성"] = agent_tbl[avg_col].apply(
                lambda v: "✅" if v >= score_threshold else "❌"
            )
        st.dataframe(agent_tbl, use_container_width=True, hide_index=True)

        # ⑤ 점수 분포
        sec("⑤ 최종점수 분포")
        f5 = fig_score_hist(df, score_col, "최종점수 분포", score_threshold)
        if f5: st.plotly_chart(f5, use_container_width=True)

    st.text_area("💬 Comment", key="t2_cmt",
                 placeholder="상담사별 특이사항을 입력하세요.")
    st.text_area("📋 Action Plan", key="t2_plan",
                 placeholder="개인별 코칭 계획을 입력하세요.")

# ══════════════════════════════════════════
# TAB 3 : 70점 미만 모니터링 ★ 핵심 장표
# ══════════════════════════════════════════
with tab3:
    sec(f"🔴 {score_threshold:.0f}점 미만 모니터링 대상")

    score_col = "최종점수" if "최종점수" in df.columns else "_score"
    below_df  = df[to_num(df[score_col]) < score_threshold].copy()

    total_cnt   = len(df)
    below_cnt   = len(below_df)
    below_pct   = below_cnt / total_cnt * 100 if total_cnt else 0
    fb_done     = (below_df["피드백여부"] == "O").sum() if "피드백여부" in below_df.columns else 0
    fb_rate     = fb_done / below_cnt * 100 if below_cnt else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.markdown(kpi("전체 건수",        total_cnt, fmt=",d"), unsafe_allow_html=True)
    c2.markdown(kpi(f"{score_threshold:.0f}점 미만 건수", below_cnt, fmt=",d",
                    sub=f"전체의 {below_pct:.1f}%"), unsafe_allow_html=True)
    c3.markdown(kpi("피드백 완료",      fb_done, fmt=",d",
                    sub=f"완료율 {fb_rate:.1f}%"), unsafe_allow_html=True)
    c4.markdown(kpi("피드백 미완료",    below_cnt-fb_done, fmt=",d"), unsafe_allow_html=True)

    if below_cnt == 0:
        alert(f"🎉 {score_threshold:.0f}점 미만 건수가 없습니다! 훌륭합니다.", "good")
    else:
        if fb_rate < 70:
            alert(f"⚠️ 피드백 완료율이 <b>{fb_rate:.1f}%</b>입니다. 미완료 건 조속 처리 필요.", "danger")

        st.markdown("---")

        ca3, cb3 = st.columns(2)
        with ca3:
            sec("① 상담사별 70점 미만 건수")
            if "상담사" in below_df.columns:
                ag_below = below_df.groupby("상담사").agg(
                    건수=(score_col,"count"),
                    평균점수=(score_col, lambda x: round(safe_mean(x),1))
                ).reset_index().sort_values("건수", ascending=False)
                fig_b = px.bar(ag_below, x="상담사", y="건수",
                               color="평균점수",
                               color_continuous_scale="Reds_r",
                               text="건수",
                               title=f"상담사별 {score_threshold:.0f}점 미만 건수")
                fig_b.update_traces(textposition="outside")
                fig_b.update_layout(**CHART_LAYOUT, height=380,
                                    showlegend=False)
                st.plotly_chart(fig_b, use_container_width=True)

        with cb3:
            sec("② 주차별 70점 미만 추이")
            if "회신주차" in below_df.columns:
                wk_below = below_df.groupby("회신주차")[score_col].count().reset_index()
                wk_below.columns = ["주차","건수"]
                def wk_sort(w):
                    try:
                        m=int(re.search(r"(\d+)월",str(w)).group(1))
                        n=int(re.search(r"(\d+)주",str(w)).group(1))
                        return m*10+n
                    except: return 0
                wk_below = wk_below.sort_values("주차", key=lambda s:s.map(wk_sort))
                fig_wk = px.bar(wk_below, x="주차", y="건수",
                                color="건수",
                                color_continuous_scale="Reds",
                                text="건수",
                                title="주차별 70점 미만 건수")
                fig_wk.update_traces(textposition="outside")
                fig_wk.update_layout(**CHART_LAYOUT, height=380,
                                     showlegend=False)
                st.plotly_chart(fig_wk, use_container_width=True)

        # ③ 피드백 현황
        sec("③ 피드백 현황")
        cc3, cd3 = st.columns(2)
        with cc3:
            if "피드백여부" in below_df.columns:
                fb_vc = below_df["피드백여부"].value_counts().reset_index()
                fb_vc.columns = ["피드백여부","건수"]
                fig_fb = px.pie(fb_vc, values="건수", names="피드백여부",
                                color="피드백여부",
                                color_discrete_map={"O":"#43a047","X":"#ef5350"},
                                hole=0.4, title="피드백 완료 여부")
                fig_fb.update_traces(textinfo="percent+label")
                fig_fb.update_layout(**CHART_LAYOUT, height=300)
                st.plotly_chart(fig_fb, use_container_width=True)

        with cd3:
            if "채널_구분" in below_df.columns:
                ch_below = below_df.groupby("채널_구분")[score_col].agg(["count","mean"]).reset_index()
                ch_below.columns = ["채널","건수","평균점수"]
                ch_below["평균점수"] = ch_below["평균점수"].round(1)
                fig_ch = px.bar(ch_below, x="채널", y="건수",
                                color="채널", text="건수",
                                color_discrete_sequence=COLORS,
                                title="채널별 70점 미만 건수")
                fig_ch.update_traces(textposition="outside")
                fig_ch.update_layout(**CHART_LAYOUT, height=300,
                                     showlegend=False)
                st.plotly_chart(fig_ch, use_container_width=True)

        # ④ 전체 목록
        sec(f"④ {score_threshold:.0f}점 미만 전체 목록")

        show_cols = []
        for c in ["회신주차","회신일자","상담사","상담사근속","채널_구분",
                  score_col,"친절점수","만족점수","이행점수",
                  "귀책분류","유형","키워드","피드백여부","피드백결과","상세분석"]:
            if c in below_df.columns:
                show_cols.append(c)

        disp = below_df[show_cols].sort_values(score_col).copy()

        # 조건부 스타일링
        def color_score(val):
            try:
                v = float(val)
                if v < 40: return "background-color:#fde8e8;color:#c0392b;font-weight:bold"
                elif v < score_threshold: return "background-color:#fff3cd"
                return ""
            except: return ""

        st.dataframe(
            disp.style.applymap(color_score, subset=[score_col] if score_col in disp.columns else []),
            use_container_width=True,
            hide_index=True,
            height=400,
        )

        # 다운로드
        csv_bytes = disp.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label=f"📥 {score_threshold:.0f}점 미만 목록 다운로드 (CSV)",
            data=csv_bytes,
            file_name=f"csat_below{score_threshold:.0f}_{report_month}.csv",
            mime="text/csv",
        )

    st.text_area("💬 Comment", key="t3_cmt",
                 placeholder="70점 미만 건 특이사항을 입력하세요.")
    st.text_area("📋 Action Plan", key="t3_plan",
                 placeholder="70점 미만 건 개선 계획을 입력하세요.")

# ══════════════════════════════════════════
# TAB 4 : 귀책 & 불만 분석
# ══════════════════════════════════════════
with tab4:
    sec("🔍 귀책 분류 & 불만 유형 분석")

    score_col = "최종점수" if "최종점수" in df.columns else "_score"

    ca4, cb4 = st.columns(2)
    with ca4:
        sec("① 귀책 분류별 건수")
        f = fig_reason_bar(df, "귀책분류", "귀책분류별 건수 (TOP 10)", top_n=10)
        if f: st.plotly_chart(f, use_container_width=True)
        else: st.info("귀책분류 컬럼 없음")

    with cb4:
        sec("② 불만 유형별 건수")
        f2 = fig_reason_bar(df, "유형", "불만 유형별 건수 (TOP 10)", top_n=10)
        if f2: st.plotly_chart(f2, use_container_width=True)
        else: st.info("유형 컬럼 없음")

    # ③ 귀책별 평균점수
    sec("③ 귀책 분류별 평균 최종점수")
    if "귀책분류" in df.columns and score_col in df.columns:
        resp = df.groupby("귀책분류")[score_col].agg(["mean","count"]).reset_index()
        resp.columns = ["귀책분류","평균점수","건수"]
        resp["평균점수"] = resp["평균점수"].round(1)
        resp = resp.sort_values("평균점수")
        colors_r = ["#ef5350" if v < score_threshold else "#3949ab" for v in resp["평균점수"]]
        fig3 = go.Figure(go.Bar(
            x=resp["평균점수"], y=resp["귀책분류"],
            orientation="h", marker_color=colors_r,
            text=resp["평균점수"], textposition="outside",
        ))
        fig3.add_vline(x=score_threshold, line_dash="dot", line_color="#ef5350")
        fig3.update_layout(**CHART_LAYOUT,
                           height=max(300, len(resp)*38),
                           xaxis=dict(range=[0,110]))
        st.plotly_chart(fig3, use_container_width=True)

    # ④ 귀책 × 채널 크로스탭
    sec("④ 귀책분류 × 채널 크로스탭")
    if "귀책분류" in df.columns:
        cross = pd.crosstab(df["귀책분류"], df["채널_구분"])
        st.dataframe(cross, use_container_width=True)

    # ⑤ 상담사별 귀책 상세표
    sec("⑤ 상담사별 귀책 & 유형 상세")
    if "상담사" in df.columns and "귀책분류" in df.columns:
        detail = df.groupby(["상담사","귀책분류"]).agg(
            건수=(score_col,"count"),
            평균점수=(score_col, lambda x: round(safe_mean(x),1)),
        ).reset_index().sort_values(["상담사","건수"], ascending=[True,False])
        st.dataframe(detail, use_container_width=True, hide_index=True)

    # ⑥ 이행점수 분포
    if "이행점수" in df.columns:
        sec("⑥ 이행점수 분포")
        f6 = fig_score_hist(df, "이행점수", "이행점수 분포", 80)
        if f6: st.plotly_chart(f6, use_container_width=True)

    st.text_area("💬 Comment", key="t4_cmt",
                 placeholder="귀책/불만 분석 특이사항을 입력하세요.")
    st.text_area("📋 Action Plan", key="t4_plan",
                 placeholder="귀책 개선 계획을 입력하세요.")

# ─────────────────────────────────────────
# 10. FOOTER
# ─────────────────────────────────────────
st.markdown("---")
st.caption(f"⭐ CSAT 품질 대시보드 · {report_month} 기준 · 모니터링 기준 {score_threshold:.0f}점 미만 · Powered by Streamlit & Plotly")