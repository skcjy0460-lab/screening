"""
병원 청구심사 월간결산 분석 대시보드
사용법: streamlit run medical_review_app.py
필요 패키지: pip install streamlit pandas openpyxl plotly
"""

import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")

# ── 페이지 설정 ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="청구심사 월간결산 대시보드",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS 스타일 ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* 전체 배경 */
    .main { background-color: #f8f9fb; }

    /* 헤더 */
    .dashboard-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2e6da4 100%);
        padding: 24px 32px;
        border-radius: 12px;
        color: white;
        margin-bottom: 24px;
    }
    .dashboard-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .dashboard-header p  { margin: 6px 0 0; font-size: 0.95rem; opacity: 0.85; }

    /* KPI 카드 */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        border-left: 5px solid #2e6da4;
        height: 100%;
    }
    .kpi-card.warning { border-left-color: #e67e22; }
    .kpi-card.danger  { border-left-color: #e74c3c; }
    .kpi-card.success { border-left-color: #27ae60; }
    .kpi-label { font-size: 0.82rem; color: #7f8c8d; font-weight: 600;
                 text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-value { font-size: 1.7rem; font-weight: 800; color: #1e3a5f; margin: 6px 0 2px; }
    .kpi-sub   { font-size: 0.82rem; color: #95a5a6; }

    /* 섹션 타이틀 */
    .section-title {
        font-size: 1.05rem; font-weight: 700; color: #1e3a5f;
        border-bottom: 2px solid #2e6da4; padding-bottom: 6px;
        margin: 28px 0 16px;
    }

    /* 알람 배지 */
    .alert-badge {
        display: inline-block;
        padding: 3px 10px; border-radius: 20px;
        font-size: 0.78rem; font-weight: 700;
    }
    .alert-high   { background:#fde8e8; color:#c0392b; }
    .alert-medium { background:#fef3cd; color:#856404; }
    .alert-low    { background:#d4edda; color:#155724; }

    /* 표 스타일 */
    .styled-table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    .styled-table th {
        background: #2e6da4; color: white;
        padding: 9px 12px; text-align: center; white-space: nowrap;
    }
    .styled-table td { padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right; }
    .styled-table tr:hover { background: #f0f4f9; }
    .styled-table .left { text-align: left; }

    /* 탭 스타일 */
    .stTabs [data-baseweb="tab"] {
        font-weight: 600; font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  데이터 파싱 함수
# ══════════════════════════════════════════════════════════════════════════════

def fmt_won(v):
    """숫자를 '1,234,567 원' 형태로 포맷"""
    try:
        return f"{int(v):,} 원"
    except Exception:
        return "-"

def fmt_pct(v, decimals=2):
    try:
        return f"{v:.{decimals}f}%"
    except Exception:
        return "-"

@st.cache_data
def parse_excel(file_bytes, sheet_name=None):
    """
    엑셀 파일을 파싱하여 요약(summary) 딕셔너리와
    상세(detail) DataFrame을 반환한다.
    """
    # bytes → BytesIO 변환 (Streamlit uploader는 bytes를 반환)
    if isinstance(file_bytes, (bytes, bytearray)):
        file_bytes = io.BytesIO(file_bytes)

    xl = pd.ExcelFile(file_bytes)
    sheets = xl.sheet_names

    if sheet_name is None:
        sheet_name = sheets[0]

    # ── 원시 데이터 전체 읽기 ──
    file_bytes.seek(0)
    raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)

    # ── 요약 테이블 (상단 8행) ──
    # 행1: 헤더(건수 / 청구금액 / 삭감액 / 심사결정금액 / 재심금액 / 불능보류건수 / 불능보류금액)
    # 행2~4: 외래 (공단/보호/자보/소계)
    # 행5~8: 입원 (공단/보호/자보/소계)
    SUMMARY_COL = {
        "보험자": 1,
        "건수": 3,
        "청구금액": 4,
        "삭감액": 7,
        "심사결정금액": 10,
        "재심금액": 13,
        "불능보류건수": 14,
        "불능보류금액": 15,
    }

    summary_rows = []
    # 외래: rows 1-4 (0-indexed), 입원: rows 5-8
    # row index 1 → 외래공단, 2 → 외래보호, 3 → 외래자보, 4 → 외래소계
    # row index 5 → 입원공단, 6 → 입원보호, 7 → 입원자보, 8 → 입원소계
    LABEL_MAP = {
        1: ("외래", "공단"), 2: ("외래", "보호"), 3: ("외래", "자보"), 4: ("외래", "소계"),
        5: ("입원", "공단"), 6: ("입원", "보호"), 7: ("입원", "자보"), 8: ("입원", "소계"),
    }

    def safe_int(v):
        try:
            return int(float(str(v).replace(",", "")))
        except Exception:
            return 0

    for ridx, (형태, 보험자) in LABEL_MAP.items():
        row = raw.iloc[ridx]
        r = {
            "형태": 형태,
            "보험자": 보험자,
            "건수": safe_int(row.iloc[3]),
            "청구금액": safe_int(row.iloc[4]),
            "삭감액": safe_int(row.iloc[7]),
            "심사결정금액": safe_int(row.iloc[10]),
            "재심금액": safe_int(row.iloc[13]),
            "불능보류건수": safe_int(row.iloc[14]),
            "불능보류금액": safe_int(row.iloc[15]),
        }
        summary_rows.append(r)

    summary = pd.DataFrame(summary_rows)

    # ── 상세 명세 DataFrame ──
    # 헤더 행 찾기: '처리상태' 가 있는 행
    header_row = None
    for i, row in raw.iterrows():
        if "처리상태" in str(row.values):
            header_row = i
            break

    detail = None
    if header_row is not None:
        cols = raw.iloc[header_row].tolist()
        # 빈 열 이름 처리
        seen = {}
        clean_cols = []
        for c in cols:
            c = str(c).strip() if pd.notna(c) else ""
            if c == "" or c == "nan":
                c = f"_col{len(clean_cols)}"
            if c in seen:
                seen[c] += 1
                c = f"{c}_{seen[c]}"
            else:
                seen[c] = 0
            clean_cols.append(c)

        data_rows = raw.iloc[header_row + 1:].copy()
        data_rows.columns = clean_cols
        data_rows = data_rows.reset_index(drop=True)

        # 처리상태가 있는 행만 (소계 행 제거)
        data_rows = data_rows[data_rows["처리상태"].notna() & (data_rows["처리상태"] != "")]

        # 핵심 열만 정리
        keep_cols = {
            "처리상태": "처리상태",
            "접수일자": "접수일자",
            "진료년월": "진료년월",
            "청구구분": "청구구분",
            "보험자": "보험자",
            "분야": "분야",
            "형태": "형태",
            "청구건수": "청구건수",
            "청구액": "청구액",
            "심사결정액": "심사결정액",
            "삭감건수": "삭감건수",
            "삭감액": "삭감액",
            "불능/보류  건수": "불능보류건수",
            "불능/보류액": "불능보류액",
            "재심금액": "재심금액",
            "비고": "비고",
        }
        existing = {k: v for k, v in keep_cols.items() if k in data_rows.columns}
        detail = data_rows[list(existing.keys())].rename(columns=existing).copy()

        # 숫자 열 변환
        num_cols = ["청구건수", "청구액", "심사결정액", "삭감건수", "삭감액",
                    "불능보류건수", "불능보류액", "재심금액"]
        for c in num_cols:
            if c in detail.columns:
                detail[c] = pd.to_numeric(detail[c], errors="coerce").fillna(0).astype(int)

        # 삭감률 계산
        detail["삭감률(%)"] = (
            detail["삭감액"] / detail["청구액"].replace(0, pd.NA) * 100
        ).round(2)

    return xl.sheet_names, sheet_name, summary, detail


# ══════════════════════════════════════════════════════════════════════════════
#  사이드바 & 파일 업로드
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/hospital.png", width=64)
    st.title("청구심사 분석")
    st.markdown("---")

    uploaded = st.file_uploader(
        "📂 엑셀 파일 업로드",
        type=["xlsx", "xls"],
        help="월간결산 엑셀 파일을 업로드하세요. 여러 달 비교 시 파일을 추가 업로드하세요.",
        accept_multiple_files=True,
    )

    st.markdown("---")
    st.markdown("#### ⚙️ 분석 설정")
    삭감률_임계값 = st.slider("삭감률 경고 기준 (%)", 0.1, 10.0, 1.5, 0.1,
                              help="이 값 이상이면 경고 표시")
    show_pending = st.checkbox("심사중 건 포함", value=False,
                               help="처리완료 + 심사중 데이터를 함께 분석")

    st.markdown("---")
    st.caption("💡 여러 파일을 동시 업로드하면 월별 비교 탭이 활성화됩니다.")


# ══════════════════════════════════════════════════════════════════════════════
#  메인 화면
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="dashboard-header">
  <h1>🏥 청구심사 월간결산 분석 대시보드</h1>
  <p>병원 개원·경영 컨설팅 | 보험청구 심사결과 비교·분석 보고서</p>
</div>
""", unsafe_allow_html=True)

if not uploaded:
    st.info("👈 왼쪽 사이드바에서 월간결산 엑셀 파일을 업로드하세요.")
    st.markdown("""
    **지원하는 엑셀 형식:**
    - 상단 요약표 (외래/입원 × 공단/보호/자보 × 건수/청구액/삭감액/심사결정액/재심금액/불능보류)
    - 하단 상세 명세 (`처리상태`, `접수일자`, `보험자`, `분야`, `형태`, `청구건수`, `청구액`, `삭감액` … 포함)

    **분석 내용:**
    1. KPI 요약 카드
    2. 보험자별 / 형태별 분석 차트
    3. 분야별 삭감 현황
    4. 불능·보류 현황
    5. 삭감률 경고 목록
    6. 월별 추이 비교 (파일 여러 개 업로드 시)
    """)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
#  파일 파싱 (첫 번째 파일 = 주 분석 대상)
# ══════════════════════════════════════════════════════════════════════════════

all_data = {}   # {파일명: (summary, detail)}

for uf in uploaded:
    try:
        file_bytes = io.BytesIO(uf.read())
        sheets, sheet_name, summary, detail = parse_excel(file_bytes, sheet_name=None)
        label = uf.name.replace(".xlsx", "").replace(".xls", "")
        all_data[label] = (summary, detail)
    except Exception as e:
        st.error(f"파일 파싱 오류 ({uf.name}): {e}")

if not all_data:
    st.stop()

# 첫 번째 파일을 주 분석 대상으로
main_label = list(all_data.keys())[0]
summary, detail = all_data[main_label]

# 처리상태 필터
if detail is not None and not show_pending:
    detail = detail[detail["처리상태"] == "처리완료"].copy()

# 소계 행
외래_소계 = summary[(summary["형태"] == "외래") & (summary["보험자"] == "소계")].iloc[0]
입원_소계 = summary[(summary["형태"] == "입원") & (summary["보험자"] == "소계")].iloc[0]
total_청구 = 외래_소계["청구금액"] + 입원_소계["청구금액"]
total_삭감 = 외래_소계["삭감액"] + 입원_소계["삭감액"]
total_결정 = 외래_소계["심사결정금액"] + 입원_소계["심사결정금액"]
total_재심 = 외래_소계["재심금액"] + 입원_소계["재심금액"]
total_불능보류금 = 외래_소계["불능보류금액"] + 입원_소계["불능보류금액"]
total_건수 = 외래_소계["건수"] + 입원_소계["건수"]
total_삭감률 = total_삭감 / total_청구 * 100 if total_청구 else 0


# ══════════════════════════════════════════════════════════════════════════════
#  탭 구성
# ══════════════════════════════════════════════════════════════════════════════

탭목록 = ["📊 종합 요약", "📋 보험자별 분석", "🔍 상세 명세", "⚠️ 삭감 경고", "📈 월별 비교"]
tabs = st.tabs(탭목록)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 : 종합 요약
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown(f"<p style='color:#7f8c8d;font-size:0.9rem'>분석 파일: <b>{main_label}</b></p>",
                unsafe_allow_html=True)

    # ── KPI 카드 ──
    col1, col2, col3, col4, col5 = st.columns(5)

    card_style = "kpi-card"
    삭감률_color = "danger" if total_삭감률 >= 삭감률_임계값 else "success"

    def kpi_html(label, value, sub, cls=""):
        return f"""<div class="kpi-card {cls}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>"""

    with col1:
        st.markdown(kpi_html("총 청구건수", f"{total_건수:,} 건",
                             f"외래 {외래_소계['건수']:,} / 입원 {입원_소계['건수']:,}"), unsafe_allow_html=True)
    with col2:
        st.markdown(kpi_html("총 청구금액", f"{total_청구/1e6:.1f}백만",
                             fmt_won(total_청구)), unsafe_allow_html=True)
    with col3:
        st.markdown(kpi_html("총 삭감액", f"{total_삭감/1e6:.2f}백만",
                             fmt_won(total_삭감), cls=삭감률_color), unsafe_allow_html=True)
    with col4:
        st.markdown(kpi_html("삭감률", fmt_pct(total_삭감률),
                             f"기준: {삭감률_임계값}% 이상 경고", cls=삭감률_color), unsafe_allow_html=True)
    with col5:
        st.markdown(kpi_html("불능·보류 금액", f"{total_불능보류금/1e6:.2f}백만",
                             fmt_won(total_불능보류금), cls="warning" if total_불능보류금 > 0 else ""),
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 요약 표 (외래/입원 분리) ──
    st.markdown('<div class="section-title">📋 심사결과 요약표</div>', unsafe_allow_html=True)

    def render_summary_table(df_sub, 형태명):
        rows = df_sub[df_sub["형태"] == 형태명].copy()
        rows["삭감액"] = pd.to_numeric(rows["삭감액"], errors='coerce')
        rows["청구금액"] = pd.to_numeric(rows["청구금액"], errors='coerce')
        display = rows[["보험자", "건수", "청구금액", "삭감액", "삭감률(%)",
                         "심사결정금액", "재심금액", "불능보류건수", "불능보류금액"]].copy()

        def fmt_row(r):
            return {
                "보험자": r["보험자"],
                "건수": f"{int(r['건수']):,}",
                "청구금액": f"{int(r['청구금액']):,}",
                "삭감액": f"{int(r['삭감액']):,}",
                "삭감률(%)": f"{r['삭감률(%)']:.2f}%" if pd.notna(r["삭감률(%)"]) else "-",
                "심사결정금액": f"{int(r['심사결정금액']):,}",
                "재심금액": f"{int(r['재심금액']):,}",
                "불능보류건수": f"{int(r['불능보류건수']):,}",
                "불능보류금액": f"{int(r['불능보류금액']):,}",
            }

        formatted = display.apply(fmt_row, axis=1, result_type="expand")
        st.markdown(f"**{형태명}**")
        st.dataframe(formatted, use_container_width=True, hide_index=True)

    col_a, col_b = st.columns(2)
    with col_a:
        render_summary_table(summary, "외래")
    with col_b:
        render_summary_table(summary, "입원")

    # ── 도넛 차트 : 외래/입원 비중 ──
    st.markdown('<div class="section-title">📊 청구금액 구성</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        # 외래/입원 비중
        labels = ["외래", "입원"]
        values = [외래_소계["청구금액"], 입원_소계["청구금액"]]
        fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.55,
                               marker_colors=["#2e6da4", "#27ae60"]))
        fig.update_layout(title="외래/입원 청구금액 비중", height=320,
                          margin=dict(t=40, b=0, l=0, r=0))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        # 보험자별 청구금액 (소계 제외)
        sub = summary[summary["보험자"] != "소계"].copy()
        sub_grp = sub.groupby("보험자")["청구금액"].sum().reset_index()
        fig2 = go.Figure(go.Pie(
            labels=sub_grp["보험자"], values=sub_grp["청구금액"],
            hole=0.55,
            marker_colors=["#2e6da4", "#e67e22", "#e74c3c"]
        ))
        fig2.update_layout(title="보험자별 청구금액 비중", height=320,
                           margin=dict(t=40, b=0, l=0, r=0))
        st.plotly_chart(fig2, use_container_width=True)

    # ── 청구 vs 삭감 막대 ──
    st.markdown('<div class="section-title">📊 청구 / 삭감 / 심사결정 비교</div>',
                unsafe_allow_html=True)

    sub2 = summary[summary["보험자"] != "소계"].copy()
    sub2["category"] = sub2["형태"] + "_" + sub2["보험자"]
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(name="청구금액", x=sub2["category"], y=sub2["청구금액"],
                          marker_color="#2e6da4", opacity=0.85))
    fig3.add_trace(go.Bar(name="심사결정금액", x=sub2["category"], y=sub2["심사결정금액"],
                          marker_color="#27ae60", opacity=0.85))
    fig3.add_trace(go.Bar(name="삭감액", x=sub2["category"], y=sub2["삭감액"],
                          marker_color="#e74c3c", opacity=0.85))
    fig3.update_layout(barmode="group", height=380, xaxis_title="구분",
                       yaxis_title="금액 (원)", legend=dict(orientation="h", y=-0.2),
                       margin=dict(t=20, b=0))
    st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 : 보험자별 분석
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    if detail is None or detail.empty:
        st.warning("상세 명세 데이터가 없습니다.")
    else:
        # ── 보험자 × 형태 × 분야 분석 ──
        st.markdown('<div class="section-title">보험자 × 형태 × 분야별 청구/삭감 현황</div>',
                    unsafe_allow_html=True)

        grp_cols = ["보험자", "형태", "분야"]
        available_grp = [c for c in grp_cols if c in detail.columns]
        grp = (detail.groupby(available_grp)
               .agg(건수=("청구건수", "sum"),
                    청구액=("청구액", "sum"),
                    삭감액=("삭감액", "sum"),
                    심사결정액=("심사결정액", "sum"),
                    재심금액=("재심금액", "sum"),
                    불능보류액=("불능보류액", "sum") if "불능보류액" in detail.columns else ("삭감액", "count"))
               .reset_index())
        grp["삭감률(%)"] = (grp["삭감액"] / grp["청구액"].replace(0, pd.NA) * 100).round(2)

        # 필터
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            선택형태 = st.multiselect("형태 필터", ["외래", "입원"], default=["외래", "입원"])
        with col_f2:
            if "보험자" in grp.columns:
                선택보험자 = st.multiselect("보험자 필터",
                                            grp["보험자"].unique().tolist(),
                                            default=grp["보험자"].unique().tolist())
            else:
                선택보험자 = []

        mask = grp["형태"].isin(선택형태)
        if 선택보험자:
            mask &= grp["보험자"].isin(선택보험자)
        grp_filtered = grp[mask]

        # 히트맵 : 분야별 삭감률
        if "분야" in grp_filtered.columns:
            pivot = grp_filtered.pivot_table(
                index="분야", columns="형태", values="삭감률(%)", aggfunc="mean"
            ).round(2)

            fig_heat = go.Figure(go.Heatmap(
                z=pivot.values,
                x=pivot.columns.tolist(),
                y=pivot.index.tolist(),
                colorscale="RdYlGn_r",
                zmin=0, zmax=10,
                text=[[f"{v:.2f}%" if pd.notna(v) else "-" for v in row] for row in pivot.values],
                texttemplate="%{text}",
            ))
            fig_heat.update_layout(title="분야 × 형태별 삭감률 히트맵 (%)",
                                   height=320, margin=dict(t=40, b=0))
            st.plotly_chart(fig_heat, use_container_width=True)

        # 상세 표
        display_cols = [c for c in ["보험자", "형태", "분야", "건수", "청구액",
                                     "삭감액", "삭감률(%)", "심사결정액", "재심금액"]
                        if c in grp_filtered.columns]
        st.dataframe(
            grp_filtered[display_cols].sort_values("삭감액", ascending=False).reset_index(drop=True),
            use_container_width=True, hide_index=True
        )

        # ── 분야별 삭감액 가로 막대 ──
        st.markdown('<div class="section-title">분야별 삭감액</div>', unsafe_allow_html=True)
        if "분야" in grp_filtered.columns:
            분야그룹 = (grp_filtered.groupby("분야")
                        .agg(삭감액=("삭감액", "sum"), 청구액=("청구액", "sum"))
                        .reset_index())
            분야그룹["삭감률(%)"] = (분야그룹["삭감액"] / 분야그룹["청구액"].replace(0, pd.NA) * 100).round(2)
            분야그룹 = 분야그룹.sort_values("삭감액")

            fig_bar = go.Figure(go.Bar(
                x=분야그룹["삭감액"],
                y=분야그룹["분야"],
                orientation="h",
                marker_color=["#e74c3c" if v >= 삭감률_임계값 else "#2e6da4"
                              for v in 분야그룹["삭감률(%)"]],
                text=[f"{v:,}원 ({r:.2f}%)" for v, r in
                      zip(분야그룹["삭감액"], 분야그룹["삭감률(%)"])],
                textposition="outside",
            ))
            fig_bar.update_layout(height=300, margin=dict(t=10, b=0, l=0, r=120),
                                  xaxis_title="삭감액 (원)")
            st.plotly_chart(fig_bar, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 : 상세 명세
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    if detail is None or detail.empty:
        st.warning("상세 명세 데이터가 없습니다.")
    else:
        st.markdown('<div class="section-title">청구 상세 명세</div>', unsafe_allow_html=True)

        # 검색·필터 옵션
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            f_보험자 = st.multiselect("보험자", detail["보험자"].dropna().unique().tolist(),
                                      default=detail["보험자"].dropna().unique().tolist(),
                                      key="det_ins")
        with col_s2:
            f_형태 = st.multiselect("형태", detail["형태"].dropna().unique().tolist(),
                                    default=detail["형태"].dropna().unique().tolist(),
                                    key="det_type")
        with col_s3:
            f_청구구분 = st.multiselect("청구구분", detail["청구구분"].dropna().unique().tolist() if "청구구분" in detail.columns else [],
                                        default=detail["청구구분"].dropna().unique().tolist() if "청구구분" in detail.columns else [],
                                        key="det_cls")

        mask2 = (detail["보험자"].isin(f_보험자)) & (detail["형태"].isin(f_형태))
        if f_청구구분 and "청구구분" in detail.columns:
            mask2 &= detail["청구구분"].isin(f_청구구분)

        det_filtered = detail[mask2].copy()

        # 삭감률 경고 강조
        def highlight_삭감(row):
            rate = row.get("삭감률(%)", 0)
            try:
                if float(rate) >= 삭감률_임계값:
                    return ["background-color: #fde8e8"] * len(row)
            except Exception:
                pass
            return [""] * len(row)

        show_cols = [c for c in ["처리상태", "접수일자", "진료년월", "청구구분",
                                  "보험자", "분야", "형태", "청구건수", "청구액",
                                  "삭감액", "삭감률(%)", "심사결정액", "재심금액",
                                  "불능보류건수", "불능보류액", "비고"]
                     if c in det_filtered.columns]

        st.dataframe(
            det_filtered[show_cols].style.apply(highlight_삭감, axis=1),
            use_container_width=True, hide_index=True, height=500
        )

        # 다운로드
        csv = det_filtered[show_cols].to_csv(index=False, encoding="utf-8-sig")
        st.download_button("⬇️ CSV 다운로드", data=csv,
                           file_name=f"{main_label}_상세명세.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 4 : 삭감 경고
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-title">⚠️ 삭감률 경고 항목</div>', unsafe_allow_html=True)
    st.caption(f"삭감률 {삭감률_임계값}% 이상인 청구 건을 표시합니다.")

    if detail is None or detail.empty:
        st.warning("상세 명세 데이터가 없습니다.")
    else:
        warn = detail[detail["삭감률(%)"] >= 삭감률_임계값].copy()
        warn = warn.sort_values("삭감률(%)", ascending=False)

        if warn.empty:
            st.success(f"✅ 삭감률 {삭감률_임계값}% 이상인 항목이 없습니다.")
        else:
            # 요약 지표
            c1, c2, c3 = st.columns(3)
            c1.metric("경고 건수", f"{len(warn)} 건")
            c2.metric("경고 건 삭감 합계", f"{warn['삭감액'].sum():,} 원")
            c3.metric("경고 건 최대 삭감률", f"{warn['삭감률(%)'].max():.2f}%")

            show_w = [c for c in ["처리상태", "보험자", "분야", "형태",
                                   "청구건수", "청구액", "삭감액", "삭감률(%)",
                                   "재심금액", "비고"]
                      if c in warn.columns]
            st.dataframe(warn[show_w].reset_index(drop=True), use_container_width=True, hide_index=True)

        # 불능·보류 현황
        st.markdown('<div class="section-title">🔒 불능·보류 현황</div>', unsafe_allow_html=True)
        if "불능보류액" in detail.columns:
            pend = detail[detail["불능보류액"] > 0].copy()
            if pend.empty:
                st.success("✅ 불능·보류 건이 없습니다.")
            else:
                c1, c2 = st.columns(2)
                c1.metric("불능·보류 건수", f"{len(pend)} 건")
                c2.metric("불능·보류 금액 합계", f"{pend['불능보류액'].sum():,} 원")

                show_p = [c for c in ["처리상태", "접수일자", "진료년월", "청구구분",
                                       "보험자", "분야", "형태", "청구건수",
                                       "청구액", "불능보류건수", "불능보류액", "비고"]
                          if c in pend.columns]
                st.dataframe(pend[show_p].reset_index(drop=True), use_container_width=True, hide_index=True)

        # 재심 현황
        st.markdown('<div class="section-title">🔄 재심 현황</div>', unsafe_allow_html=True)
        if "재심금액" in detail.columns:
            resv = detail[detail["재심금액"] > 0].copy()
            if resv.empty:
                st.info("재심 청구 건이 없습니다.")
            else:
                c1, c2 = st.columns(2)
                c1.metric("재심 청구 건수", f"{len(resv)} 건")
                c2.metric("재심 금액 합계", f"{resv['재심금액'].sum():,} 원")

                show_r = [c for c in ["처리상태", "보험자", "분야", "형태",
                                       "청구액", "삭감액", "재심금액", "비고"]
                          if c in resv.columns]
                st.dataframe(resv[show_r].reset_index(drop=True), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 5 : 월별 비교
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    if len(all_data) < 2:
        st.info("👈 **월별 비교**를 보려면 사이드바에서 2개 이상의 파일을 업로드하세요.")
    else:
        st.markdown('<div class="section-title">📈 월별 추이 비교</div>', unsafe_allow_html=True)

        rows_trend = []
        for label, (s, d) in all_data.items():
            외 = s[(s["형태"] == "외래") & (s["보험자"] == "소계")].iloc[0]
            입 = s[(s["형태"] == "입원") & (s["보험자"] == "소계")].iloc[0]
            tc = 외["청구금액"] + 입["청구금액"]
            td = 외["삭감액"] + 입["삭감액"]
            rows_trend.append({
                "월": label,
                "총청구금액": tc,
                "총삭감액": td,
                "삭감률(%)": round(td / tc * 100, 2) if tc else 0,
                "외래건수": 외["건수"],
                "입원건수": 입["건수"],
                "재심금액": 외["재심금액"] + 입["재심금액"],
                "불능보류금액": 외["불능보류금액"] + 입["불능보류금액"],
            })

        trend = pd.DataFrame(rows_trend)

        fig_trend = make_subplots(rows=2, cols=2,
                                  subplot_titles=["총 청구금액 추이", "삭감률(%) 추이",
                                                  "외래/입원 건수", "재심·불능보류 금액"])

        fig_trend.add_trace(go.Scatter(x=trend["월"], y=trend["총청구금액"],
                                       mode="lines+markers", name="청구금액",
                                       line=dict(color="#2e6da4", width=2)), row=1, col=1)
        fig_trend.add_trace(go.Scatter(x=trend["월"], y=trend["총삭감액"],
                                       mode="lines+markers", name="삭감액",
                                       line=dict(color="#e74c3c", width=2)), row=1, col=1)

        fig_trend.add_trace(go.Scatter(x=trend["월"], y=trend["삭감률(%)"],
                                       mode="lines+markers+text",
                                       text=[f"{v:.2f}%" for v in trend["삭감률(%)"]],
                                       textposition="top center",
                                       name="삭감률",
                                       line=dict(color="#e67e22", width=2)), row=1, col=2)
        fig_trend.add_hline(y=삭감률_임계값, line_dash="dash", line_color="red",
                            annotation_text=f"기준 {삭감률_임계값}%", row=1, col=2)

        fig_trend.add_trace(go.Bar(x=trend["월"], y=trend["외래건수"],
                                   name="외래건수", marker_color="#2e6da4"), row=2, col=1)
        fig_trend.add_trace(go.Bar(x=trend["월"], y=trend["입원건수"],
                                   name="입원건수", marker_color="#27ae60"), row=2, col=1)

        fig_trend.add_trace(go.Bar(x=trend["월"], y=trend["재심금액"],
                                   name="재심금액", marker_color="#8e44ad"), row=2, col=2)
        fig_trend.add_trace(go.Bar(x=trend["월"], y=trend["불능보류금액"],
                                   name="불능보류금액", marker_color="#e67e22"), row=2, col=2)

        fig_trend.update_layout(height=640, barmode="group",
                                 legend=dict(orientation="h", y=-0.08),
                                 margin=dict(t=40, b=0))
        st.plotly_chart(fig_trend, use_container_width=True)

        st.dataframe(trend, use_container_width=True, hide_index=True)


# ── 푸터 ──────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#b0b8c1;font-size:0.8rem'>"
    "🏥 병원 개원·경영 컨설팅 청구심사 분석 대시보드 | "
    "Powered by Streamlit + Plotly"
    "</div>",
    unsafe_allow_html=True,
)
