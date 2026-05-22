"""
병원 청구심사 월간결산 분석 대시보드
사용법: streamlit run medical_review_app.py
필요 패키지: pip install streamlit pandas openpyxl plotly
"""

import io
import re
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")

# ── 페이지 설정 ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="청구심사 종합 분석 대시보드",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS 스타일 ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fb; }
    .dashboard-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2e6da4 100%);
        padding: 24px 32px; border-radius: 12px;
        color: white; margin-bottom: 24px;
    }
    .dashboard-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .dashboard-header p  { margin: 6px 0 0; font-size: 0.95rem; opacity: 0.85; }
    .kpi-card {
        background: white; border-radius: 12px; padding: 20px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        border-left: 5px solid #2e6da4; height: 100%;
    }
    .kpi-card.warning { border-left-color: #e67e22; }
    .kpi-card.danger  { border-left-color: #e74c3c; }
    .kpi-card.success { border-left-color: #27ae60; }
    .kpi-label { font-size: 0.82rem; color: #7f8c8d; font-weight: 600;
                 text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-value { font-size: 1.7rem; font-weight: 800; color: #1e3a5f; margin: 6px 0 2px; }
    .kpi-sub   { font-size: 0.82rem; color: #95a5a6; }
    .section-title {
        font-size: 1.05rem; font-weight: 700; color: #1e3a5f;
        border-bottom: 2px solid #2e6da4; padding-bottom: 6px; margin: 28px 0 16px;
    }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  유틸 함수
# ══════════════════════════════════════════════════════════════════════════════
def fmt_won(v):
    try:
        return f"{int(v):,} 원"
    except Exception:
        return "-"

def fmt_pct(v, d=2):
    try:
        return f"{v:.{d}f}%"
    except Exception:
        return "-"

def fmt_acct(v):
    """회계 서식: 양수=1,234,567 / 0=- / 음수=(1,234,567)"""
    try:
        n = int(v)
        if n == 0:
            return "-"
        if n < 0:
            return f"({abs(n):,})"
        return f"{n:,}"
    except Exception:
        return "-"

def parse_date(v):
    """접수일자를 YYYY-MM-DD 로 통일"""
    s = str(v).strip()
    # 이미 날짜 형식 2025-07-02 00:00:00
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    # 8자리 숫자 20250702
    m2 = re.match(r"(\d{4})(\d{2})(\d{2})", s)
    if m2:
        return f"{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"
    return s if s not in ("", "nan") else ""


# ══════════════════════════════════════════════════════════════════════════════
#  데이터 파싱 (원본 서식 완전 지원)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def parse_excel(file_bytes, sheet_name=None):
    if isinstance(file_bytes, (bytes, bytearray)):
        file_bytes = io.BytesIO(file_bytes)
    file_bytes.seek(0)

    xl = pd.ExcelFile(file_bytes)
    sheets = xl.sheet_names
    data_sheets = [s for s in sheets if not str(s).startswith("📋")]
    if sheet_name is None:
        sheet_name = data_sheets[0] if data_sheets else sheets[0]

    file_bytes.seek(0)
    raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None, dtype=str)
    raw = raw.fillna("")

    # ── 헤더 행 찾기 ──────────────────────────────────────────────────────
    header_row = None
    for i, row in raw.iterrows():
        if "처리상태" in row.values:
            header_row = i
            break

    detail = None
    if header_row is not None:
        # 열 이름 정리
        raw_cols = raw.iloc[header_row].tolist()
        seen = {}
        clean_cols = []
        for c in raw_cols:
            c = str(c).replace("\n", " ").replace("\r", "").strip()
            if c in ("", "nan"):
                c = f"_col{len(clean_cols)}"
            key = c
            if key in seen:
                seen[key] += 1
                c = f"{c}_{seen[key]}"
            else:
                seen[key] = 0
            clean_cols.append(c)

        data_rows = raw.iloc[header_row + 1:].copy()
        data_rows.columns = clean_cols
        data_rows = data_rows.reset_index(drop=True)

        # ── 열 이름 매핑 ──────────────────────────────────────────────────
        target_keys = {
            "처리상태": "처리상태",
            "접수일자": "접수일자",
            "진료년월": "진료년월",
            "접수번호": "접수번호",
            "청구구분": "청구구분",
            "보험자": "보험자",
            "분야": "분야",
            "형태": "형태",
            "청구건수": "청구건수",
            "청구액": "청구액",
            "요양급여 비용총액": "요양급여비용총액",
            "심사결정액": "심사결정액",
            "삭감건수": "삭감건수",
            "삭감액": "삭감액",
            "불능/보류 건수": "불능보류건수",
            "불능/보류액": "불능보류액",
            "재심금액": "재심금액",
            "비고": "비고",
        }
        col_map = {}
        for col in data_rows.columns:
            col_clean = col.replace("\n", " ").strip()
            if col_clean in target_keys:
                col_map[col] = target_keys[col_clean]
        data_rows = data_rows.rename(columns=col_map)

        # ── 핵심: 원본 서식의 "위에서 내려오는 값" 채우기 ────────────────
        # 원본 서식에서는 보험자·분야·형태가 첫 행에만 있고 아래 행은 비어 있음
        # 처리상태·보험자·분야·형태를 위 행 값으로 forward-fill
        fill_cols = ["보험자", "분야", "형태"]
        for col in fill_cols:
            if col in data_rows.columns:
                # 빈 문자열을 NaN으로 바꾼 뒤 ffill
                data_rows[col] = data_rows[col].replace("", pd.NA).ffill()
                data_rows[col] = data_rows[col].fillna("")

        # ── 유효 행 판별 ──────────────────────────────────────────────────
        # 규칙: 청구건수가 숫자이고 청구액이 숫자인 행 (소계행 포함 가능성 있어 추가 필터)
        # 소계행 특징: 처리상태가 비어있고 접수번호도 비어있음
        # → 처리상태가 '처리완료' 또는 '심사중' 인 행만 데이터 행으로 인정
        if "처리상태" in data_rows.columns:
            data_rows = data_rows[
                data_rows["처리상태"].isin(["처리완료", "심사중"])
            ].copy()

        # 청구액이 실제 숫자인 행만
        if "청구액" in data_rows.columns:
            data_rows = data_rows[
                pd.to_numeric(data_rows["청구액"], errors="coerce").notna()
            ].copy()

        # ── 숫자 열 변환 ──────────────────────────────────────────────────
        num_cols = [
            "청구건수", "청구액", "요양급여비용총액", "심사결정액",
            "삭감건수", "삭감액", "불능보류건수", "불능보류액", "재심금액",
        ]
        for c in num_cols:
            if c in data_rows.columns:
                data_rows[c] = (
                    pd.to_numeric(data_rows[c], errors="coerce").fillna(0).astype(int)
                )

        # ── 접수일자 통일 ─────────────────────────────────────────────────
        if "접수일자" in data_rows.columns:
            data_rows["접수일자"] = data_rows["접수일자"].apply(parse_date)

        # ── 삭감률 계산 ───────────────────────────────────────────────────
        data_rows["삭감률(%)"] = 0.0
        if "청구액" in data_rows.columns and "삭감액" in data_rows.columns:
            mask = data_rows["청구액"] > 0
            data_rows.loc[mask, "삭감률(%)"] = (
                data_rows.loc[mask, "삭감액"] /
                data_rows.loc[mask, "청구액"] * 100
            ).round(2)

        detail = data_rows.reset_index(drop=True)

    # ── 요약표: 상세에서 자동 집계 ───────────────────────────────────────
    summary_rows = []
    if detail is not None and not detail.empty:
        for 형태 in ["외래", "입원"]:
            for 보험자 in ["공단", "보호", "자보"]:
                sub = detail[
                    (detail["형태"] == 형태) & (detail["보험자"] == 보험자)
                ]
                def s(col):
                    return int(sub[col].sum()) if col in sub.columns else 0
                summary_rows.append({
                    "형태": 형태, "보험자": 보험자,
                    "건수": s("청구건수"),
                    "청구금액": s("청구액"),
                    "삭감액": s("삭감액"),
                    "심사결정금액": s("심사결정액"),
                    "재심금액": s("재심금액"),
                    "불능보류건수": s("불능보류건수"),
                    "불능보류금액": s("불능보류액"),
                })
        for 형태 in ["외래", "입원"]:
            sub_rows = [r for r in summary_rows if r["형태"] == 형태]
            summary_rows.append({
                "형태": 형태, "보험자": "소계",
                "건수":         sum(r["건수"] for r in sub_rows),
                "청구금액":     sum(r["청구금액"] for r in sub_rows),
                "삭감액":       sum(r["삭감액"] for r in sub_rows),
                "심사결정금액": sum(r["심사결정금액"] for r in sub_rows),
                "재심금액":     sum(r["재심금액"] for r in sub_rows),
                "불능보류건수": sum(r["불능보류건수"] for r in sub_rows),
                "불능보류금액": sum(r["불능보류금액"] for r in sub_rows),
            })

    summary = pd.DataFrame(summary_rows) if summary_rows else pd.DataFrame()
    return sheets, sheet_name, summary, detail


# ══════════════════════════════════════════════════════════════════════════════
#  사이드바
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
#  헤더
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="dashboard-header">
  <h1>🏥 청구심사 종합 분석 대시보드
    <span style="font-size:0.65rem;font-weight:400;opacity:0.75;margin-left:18px;vertical-align:middle;">
      제작자 : 주식회사 메디엄 조정윤
    </span>
  </h1>
  <p>병원 개원·경영 컨설팅 | 보험청구 심사결과 비교·분석 보고서</p>
</div>
""", unsafe_allow_html=True)

if not uploaded:
    st.info("👈 왼쪽 사이드바에서 월간결산 엑셀 파일을 업로드하세요.")
    with st.expander("📌 엑셀 서식 작성 규칙 (펼쳐보기)"):
        st.markdown("""
**✅ 데이터 행 인식 조건**
- `처리상태` 열 값이 반드시 **`처리완료`** 또는 **`심사중`** 이어야 합니다
- 소계행(처리상태 없이 숫자만 있는 행)은 자동 제외됩니다

**✅ 보험자 / 분야 / 형태 작성법**
- 원본 서식처럼 첫 행에만 값을 쓰고 아래 행은 비워도 자동으로 채워집니다
- 단, 그룹이 바뀔 때는 반드시 새 값을 입력하세요

**✅ 접수일자**
- `20250702` 또는 `2025-07-02` 어떤 형식이든 자동 변환됩니다

**✅ 숫자 열**
- 콤마·원 표시 없이 정수만 입력하세요 (예: `448400`)
- 빈 칸은 `0`으로 처리됩니다

**✅ 시트명**
- `YYYY-MM` 형식 권장 (예: `2025-07`)

**✅ 금액 단위**
- 원(KRW) 기준 정수
        """)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
#  파일 파싱
# ══════════════════════════════════════════════════════════════════════════════
all_data = {}
for uf in uploaded:
    try:
        file_bytes = io.BytesIO(uf.read())
        sheets, sheet_name, summary, detail = parse_excel(file_bytes)
        label = uf.name.replace(".xlsx", "").replace(".xls", "")
        all_data[label] = (summary, detail)
    except Exception as e:
        st.error(f"파일 파싱 오류 ({uf.name}): {e}")

if not all_data:
    st.stop()

main_label = list(all_data.keys())[0]
summary, detail = all_data[main_label]

if detail is not None and not show_pending:
    detail = detail[detail["처리상태"] == "처리완료"].copy()

def get_소계(형태):
    if summary.empty:
        return {k: 0 for k in ["청구금액","삭감액","심사결정금액","재심금액","불능보류금액","건수"]}
    row = summary[(summary["형태"] == 형태) & (summary["보험자"] == "소계")]
    if row.empty:
        return {k: 0 for k in ["청구금액","삭감액","심사결정금액","재심금액","불능보류금액","건수"]}
    return row.iloc[0].to_dict()

외래_소계 = get_소계("외래")
입원_소계 = get_소계("입원")
total_청구 = 외래_소계["청구금액"] + 입원_소계["청구금액"]
total_삭감 = 외래_소계["삭감액"]   + 입원_소계["삭감액"]
total_결정 = 외래_소계["심사결정금액"] + 입원_소계["심사결정금액"]
total_재심 = 외래_소계["재심금액"] + 입원_소계["재심금액"]
total_불능 = 외래_소계["불능보류금액"] + 입원_소계["불능보류금액"]
total_건수 = 외래_소계["건수"] + 입원_소계["건수"]
total_삭감률 = (total_삭감 / total_청구 * 100) if total_청구 else 0


# ══════════════════════════════════════════════════════════════════════════════
#  탭 구성
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs(["📊 종합 요약", "📋 보험자별 분석", "🔍 상세내역", "⚠️ 삭감 경고", "📈 월별 비교"])


# ──────────────────────────────────────────────────────────────────────────────
#  TAB 1 : 종합 요약
# ──────────────────────────────────────────────────────────────────────────────
with tabs[0]:
    st.markdown(
        f"<p style='color:#7f8c8d;font-size:0.9rem'>분석 파일: <b>{main_label}</b> "
        f"| 총 <b>{len(detail) if detail is not None else 0}</b>건 파싱됨</p>",
        unsafe_allow_html=True
    )

    def kpi(label, value, sub, cls=""):
        return (f'<div class="kpi-card {cls}">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}</div>'
                f'<div class="kpi-sub">{sub}</div></div>')

    삭감_cls = "danger" if total_삭감률 >= 삭감률_임계값 else "success"
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(kpi("총 청구건수", f"{int(total_건수):,} 건",
                        f"외래 {int(외래_소계['건수']):,} / 입원 {int(입원_소계['건수']):,}"),
                    unsafe_allow_html=True)
    with c2:
        st.markdown(kpi("총 청구금액", f"{total_청구/1e6:.1f}백만",
                        fmt_won(total_청구)), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi("총 삭감액", f"{total_삭감/1e6:.2f}백만",
                        fmt_won(total_삭감), cls=삭감_cls), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi("삭감률", fmt_pct(total_삭감률),
                        f"기준 {삭감률_임계값}% 이상 경고", cls=삭감_cls),
                    unsafe_allow_html=True)
    with c5:
        st.markdown(kpi("불능·보류 금액", f"{total_불능/1e6:.2f}백만",
                        fmt_won(total_불능),
                        cls="warning" if total_불능 > 0 else ""),
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">📋 심사결과 요약표</div>', unsafe_allow_html=True)

    def render_summary(형태명):
        rows = summary[summary["형태"] == 형태명].copy()
        rows["삭감률(%)"] = rows.apply(
            lambda r: round(r["삭감액"] / r["청구금액"] * 100, 2) if r["청구금액"] > 0 else 0.0,
            axis=1
        )
        disp = rows[["보험자","건수","청구금액","삭감액","삭감률(%)",
                     "심사결정금액","재심금액","불능보류건수","불능보류금액"]].copy()
        for col in ["청구금액","삭감액","심사결정금액","재심금액","불능보류금액"]:
            disp[col] = disp[col].apply(fmt_acct)
        disp["건수"]       = disp["건수"].apply(lambda v: f"{int(v):,}")
        disp["삭감률(%)"]  = disp["삭감률(%)"].apply(lambda v: f"{v:.2f}%")
        disp["불능보류건수"] = disp["불능보류건수"].apply(lambda v: f"{int(v):,}")
        st.markdown(f"**{형태명}**")
        st.dataframe(disp, use_container_width=True, hide_index=True)

    ca, cb = st.columns(2)
    with ca:
        render_summary("외래")
    with cb:
        render_summary("입원")

    st.markdown('<div class="section-title">📊 청구금액 구성</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure(go.Pie(
            labels=["외래","입원"],
            values=[외래_소계["청구금액"], 입원_소계["청구금액"]],
            hole=0.55, marker_colors=["#2e6da4","#27ae60"]
        ))
        fig.update_layout(title="외래/입원 청구금액 비중", height=320,
                          margin=dict(t=40,b=0,l=0,r=0))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        sub_s = summary[summary["보험자"] != "소계"]
        grp   = sub_s.groupby("보험자")["청구금액"].sum().reset_index()
        fig2  = go.Figure(go.Pie(
            labels=grp["보험자"], values=grp["청구금액"],
            hole=0.55, marker_colors=["#2e6da4","#e67e22","#e74c3c"]
        ))
        fig2.update_layout(title="보험자별 청구금액 비중", height=320,
                           margin=dict(t=40,b=0,l=0,r=0))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="section-title">📊 청구 / 삭감 / 심사결정 비교</div>',
                unsafe_allow_html=True)
    sub2 = summary[summary["보험자"] != "소계"].copy()
    sub2["구분"] = sub2["형태"] + "_" + sub2["보험자"]
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(name="청구금액",     x=sub2["구분"], y=sub2["청구금액"],     marker_color="#2e6da4"))
    fig3.add_trace(go.Bar(name="심사결정금액", x=sub2["구분"], y=sub2["심사결정금액"], marker_color="#27ae60"))
    fig3.add_trace(go.Bar(name="삭감액",       x=sub2["구분"], y=sub2["삭감액"],       marker_color="#e74c3c"))
    fig3.update_layout(barmode="group", height=360,
                       legend=dict(orientation="h", y=-0.2),
                       margin=dict(t=20,b=0))
    st.plotly_chart(fig3, use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
#  TAB 2 : 보험자별 분석 (개선)
# ──────────────────────────────────────────────────────────────────────────────
with tabs[1]:
    if detail is None or detail.empty:
        st.warning("상세 명세 데이터가 없습니다.")
    else:
        # ── 필터 ──────────────────────────────────────────────────────────
        cf1, cf2, cf3 = st.columns(3)
        with cf1:
            sel_형태 = st.multiselect("형태 필터", ["외래","입원"],
                                      default=["외래","입원"], key="t2_type")
        with cf2:
            ins_list = sorted(detail["보험자"].dropna().unique().tolist())
            sel_ins  = st.multiselect("보험자 필터", ins_list, default=ins_list, key="t2_ins")
        with cf3:
            분야_list = sorted(detail["분야"].dropna().unique().tolist()) if "분야" in detail.columns else []
            sel_분야  = st.multiselect("분야 필터", 분야_list, default=분야_list, key="t2_분야")

        mask = detail["형태"].isin(sel_형태) & detail["보험자"].isin(sel_ins)
        if sel_분야 and "분야" in detail.columns:
            mask &= detail["분야"].isin(sel_분야)
        det_f2 = detail[mask].copy()

        if det_f2.empty:
            st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
        else:
            # ── ① 보험자 × 형태 요약 테이블 ─────────────────────────────
            st.markdown('<div class="section-title">① 보험자 × 형태별 요약</div>',
                        unsafe_allow_html=True)

            grp_ins = (det_f2.groupby(["보험자","형태"])
                       .agg(청구건수=("청구건수","sum"),
                            청구액=("청구액","sum"),
                            삭감액=("삭감액","sum"),
                            심사결정액=("심사결정액","sum"),
                            재심금액=("재심금액","sum"))
                       .reset_index())
            grp_ins["삭감률(%)"] = grp_ins.apply(
                lambda r: round(r["삭감액"]/r["청구액"]*100,2) if r["청구액"]>0 else 0.0, axis=1)

            # 가로 막대: 보험자×형태 청구/삭감
            fig_ins = go.Figure()
            colors = {"공단":"#2e6da4","보호":"#27ae60","자보":"#e67e22"}
            shapes = {"외래":"circle","입원":"square"}
            for _, row in grp_ins.iterrows():
                label_name = f"{row['보험자']} ({row['형태']})"
                fig_ins.add_trace(go.Bar(
                    name=label_name,
                    x=["청구액","삭감액","심사결정액"],
                    y=[row["청구액"], row["삭감액"], row["심사결정액"]],
                    text=[fmt_acct(row["청구액"]), fmt_acct(row["삭감액"]), fmt_acct(row["심사결정액"])],
                    textposition="outside",
                ))
            fig_ins.update_layout(barmode="group", height=380,
                                  legend=dict(orientation="h", y=-0.25),
                                  margin=dict(t=20,b=0),
                                  yaxis_title="금액 (원)")
            st.plotly_chart(fig_ins, use_container_width=True)

            # 삭감률 테이블
            disp_ins = grp_ins.copy()
            for col in ["청구액","삭감액","심사결정액","재심금액"]:
                disp_ins[col] = disp_ins[col].apply(fmt_acct)
            disp_ins["청구건수"] = disp_ins["청구건수"].apply(lambda v: f"{int(v):,}")
            disp_ins["삭감률(%)"] = disp_ins["삭감률(%)"].apply(lambda v: f"{v:.2f}%")
            st.dataframe(disp_ins, use_container_width=True, hide_index=True)

            # ── ② 분야별 상세 ────────────────────────────────────────────
            if "분야" in det_f2.columns:
                st.markdown('<div class="section-title">② 분야별 청구 / 삭감 현황</div>',
                            unsafe_allow_html=True)

                grp_분야 = (det_f2.groupby(["분야","형태"])
                            .agg(청구건수=("청구건수","sum"),
                                 청구액=("청구액","sum"),
                                 삭감액=("삭감액","sum"),
                                 심사결정액=("심사결정액","sum"))
                            .reset_index())
                grp_분야["삭감률(%)"] = grp_분야.apply(
                    lambda r: round(r["삭감액"]/r["청구액"]*100,2) if r["청구액"]>0 else 0.0, axis=1)

                # 왼: 분야별 청구액 누적 막대 / 오: 분야별 삭감률
                d1, d2 = st.columns(2)
                with d1:
                    fig_분야_bar = go.Figure()
                    color_map = {"외래":"#2e6da4","입원":"#27ae60"}
                    for 형태 in grp_분야["형태"].unique():
                        sub_t = grp_분야[grp_분야["형태"]==형태]
                        fig_분야_bar.add_trace(go.Bar(
                            name=형태,
                            y=sub_t["분야"],
                            x=sub_t["청구액"],
                            orientation="h",
                            marker_color=color_map.get(형태,"#888"),
                        ))
                    fig_분야_bar.update_layout(
                        barmode="stack", title="분야별 청구액",
                        height=320, margin=dict(t=40,b=0,l=0,r=20),
                        legend=dict(orientation="h", y=-0.2)
                    )
                    st.plotly_chart(fig_분야_bar, use_container_width=True)

                with d2:
                    grp_분야_tot = (det_f2.groupby("분야")
                                    .agg(청구액=("청구액","sum"), 삭감액=("삭감액","sum"))
                                    .reset_index())
                    grp_분야_tot["삭감률(%)"] = grp_분야_tot.apply(
                        lambda r: round(r["삭감액"]/r["청구액"]*100,2) if r["청구액"]>0 else 0.0, axis=1)
                    grp_분야_tot = grp_분야_tot.sort_values("삭감률(%)")

                    bar_colors = ["#e74c3c" if v >= 삭감률_임계값 else "#2e6da4"
                                  for v in grp_분야_tot["삭감률(%)"]]
                    fig_분야_rate = go.Figure(go.Bar(
                        y=grp_분야_tot["분야"],
                        x=grp_분야_tot["삭감률(%)"],
                        orientation="h",
                        marker_color=bar_colors,
                        text=[f"{v:.2f}%" for v in grp_분야_tot["삭감률(%)"]],
                        textposition="outside",
                    ))
                    fig_분야_rate.update_layout(
                        title="분야별 삭감률 (%)",
                        height=320, margin=dict(t=40,b=0,l=0,r=60),
                        xaxis_title="삭감률 (%)"
                    )
                    st.plotly_chart(fig_분야_rate, use_container_width=True)

                # 분야 상세 테이블
                disp_분야 = grp_분야.copy()
                for col in ["청구액","삭감액","심사결정액"]:
                    disp_분야[col] = disp_분야[col].apply(fmt_acct)
                disp_분야["청구건수"]  = disp_분야["청구건수"].apply(lambda v: f"{int(v):,}")
                disp_분야["삭감률(%)"] = disp_분야["삭감률(%)"].apply(lambda v: f"{v:.2f}%")
                st.dataframe(disp_분야.sort_values("삭감액", ascending=False)
                             .reset_index(drop=True),
                             use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
#  TAB 3 : 상세내역
# ──────────────────────────────────────────────────────────────────────────────
with tabs[2]:
    if detail is None or detail.empty:
        st.warning("상세내역 데이터가 없습니다.")
    else:
        st.markdown('<div class="section-title">청구 상세내역</div>', unsafe_allow_html=True)

        cs1, cs2, cs3 = st.columns(3)
        with cs1:
            f_ins  = st.multiselect("보험자", detail["보험자"].dropna().unique().tolist(),
                                    default=detail["보험자"].dropna().unique().tolist(), key="d_ins")
        with cs2:
            f_type = st.multiselect("형태", detail["형태"].dropna().unique().tolist(),
                                    default=detail["형태"].dropna().unique().tolist(), key="d_type")
        with cs3:
            if "청구구분" in detail.columns:
                f_cls = st.multiselect("청구구분",
                                       detail["청구구분"].dropna().unique().tolist(),
                                       default=detail["청구구분"].dropna().unique().tolist(),
                                       key="d_cls")
            else:
                f_cls = []

        mask2 = detail["보험자"].isin(f_ins) & detail["형태"].isin(f_type)
        if f_cls and "청구구분" in detail.columns:
            mask2 &= detail["청구구분"].isin(f_cls)
        det_show = detail[mask2].copy()

        # 회계서식 적용 표시용 복사
        disp_det = det_show.copy()
        for col in ["청구액","삭감액","심사결정액","재심금액","불능보류액"]:
            if col in disp_det.columns:
                disp_det[col] = disp_det[col].apply(fmt_acct)
        # 삭감률 소수점 2자리
        if "삭감률(%)" in disp_det.columns:
            disp_det["삭감률(%)"] = disp_det["삭감률(%)"].apply(
                lambda v: f"{round(float(v), 2):.2f}%" if str(v) not in ("","nan") else "-"
            )

        def highlight_삭감(row):
            try:
                rate_str = str(row.get("삭감률(%)", "0")).replace("%","")
                if float(rate_str) >= 삭감률_임계값:
                    return ["background-color: #fde8e8"] * len(row)
            except Exception:
                pass
            return [""] * len(row)

        show_d = [c for c in ["처리상태","접수일자","진료년월","청구구분",
                               "보험자","분야","형태","청구건수","청구액",
                               "삭감액","삭감률(%)","심사결정액","재심금액",
                               "불능보류건수","불능보류액","비고"]
                  if c in disp_det.columns]
        st.dataframe(disp_det[show_d].style.apply(highlight_삭감, axis=1),
                     use_container_width=True, hide_index=True, height=520)

        st.caption(f"총 {len(det_show):,}건 표시 중 | 삭감률 {삭감률_임계값}% 이상 행은 붉은색 강조")

        csv = det_show[show_d if all(c in det_show.columns for c in show_d) else
                       [c for c in show_d if c in det_show.columns]
                      ].to_csv(index=False, encoding="utf-8-sig")
        st.download_button("⬇️ CSV 다운로드", data=csv,
                           file_name=f"{main_label}_상세내역.csv", mime="text/csv")


# ──────────────────────────────────────────────────────────────────────────────
#  TAB 4 : 삭감 경고
# ──────────────────────────────────────────────────────────────────────────────
with tabs[3]:
    st.markdown('<div class="section-title">⚠️ 삭감률 경고 항목</div>', unsafe_allow_html=True)
    st.caption(f"삭감률 {삭감률_임계값}% 이상인 청구 건을 표시합니다.")

    if detail is None or detail.empty:
        st.warning("상세내역 데이터가 없습니다.")
    else:
        warn = detail[detail["삭감률(%)"] >= 삭감률_임계값].sort_values("삭감률(%)", ascending=False)

        if warn.empty:
            st.success(f"✅ 삭감률 {삭감률_임계값}% 이상인 항목이 없습니다.")
        else:
            w1, w2, w3 = st.columns(3)
            w1.metric("경고 건수", f"{len(warn)} 건")
            w2.metric("삭감 합계", fmt_acct(warn["삭감액"].sum()) + " 원")
            w3.metric("최대 삭감률", f"{warn['삭감률(%)'].max():.2f}%")

            disp_warn = warn.copy()
            for col in ["청구액","삭감액","재심금액"]:
                if col in disp_warn.columns:
                    disp_warn[col] = disp_warn[col].apply(fmt_acct)
            if "삭감률(%)" in disp_warn.columns:
                disp_warn["삭감률(%)"] = disp_warn["삭감률(%)"].apply(
                    lambda v: f"{round(float(v),2):.2f}%")

            show_w = [c for c in ["처리상태","보험자","분야","형태",
                                   "청구건수","청구액","삭감액","삭감률(%)",
                                   "재심금액","비고"] if c in disp_warn.columns]
            st.dataframe(disp_warn[show_w].reset_index(drop=True),
                         use_container_width=True, hide_index=True)

        # 불능·보류
        st.markdown('<div class="section-title">🔒 불능·보류 현황</div>', unsafe_allow_html=True)
        if "불능보류액" in detail.columns:
            pend = detail[detail["불능보류액"] > 0]
            if pend.empty:
                st.success("✅ 불능·보류 건이 없습니다.")
            else:
                p1, p2 = st.columns(2)
                p1.metric("불능·보류 건수", f"{len(pend)} 건")
                p2.metric("불능·보류 금액", fmt_acct(pend["불능보류액"].sum()) + " 원")
                disp_pend = pend.copy()
                disp_pend["불능보류액"] = disp_pend["불능보류액"].apply(fmt_acct)
                disp_pend["청구액"]     = disp_pend["청구액"].apply(fmt_acct)
                show_p = [c for c in ["처리상태","접수일자","진료년월","청구구분",
                                       "보험자","분야","형태","청구건수",
                                       "청구액","불능보류건수","불능보류액","비고"]
                          if c in disp_pend.columns]
                st.dataframe(disp_pend[show_p].reset_index(drop=True),
                             use_container_width=True, hide_index=True)

        # 재심
        st.markdown('<div class="section-title">🔄 재심 현황</div>', unsafe_allow_html=True)
        if "재심금액" in detail.columns:
            resv = detail[detail["재심금액"] > 0]
            if resv.empty:
                st.info("재심 청구 건이 없습니다.")
            else:
                r1, r2 = st.columns(2)
                r1.metric("재심 청구 건수", f"{len(resv)} 건")
                r2.metric("재심 금액 합계", fmt_acct(resv["재심금액"].sum()) + " 원")
                disp_resv = resv.copy()
                for col in ["청구액","삭감액","재심금액"]:
                    if col in disp_resv.columns:
                        disp_resv[col] = disp_resv[col].apply(fmt_acct)
                show_r = [c for c in ["처리상태","보험자","분야","형태",
                                       "청구액","삭감액","재심금액","비고"]
                          if c in disp_resv.columns]
                st.dataframe(disp_resv[show_r].reset_index(drop=True),
                             use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
#  TAB 5 : 월별 비교
# ──────────────────────────────────────────────────────────────────────────────
with tabs[4]:
    if len(all_data) < 2:
        st.info("👈 **월별 비교**를 보려면 사이드바에서 2개 이상의 파일을 업로드하세요.")
    else:
        st.markdown('<div class="section-title">📈 월별 추이 비교</div>', unsafe_allow_html=True)
        rows_trend = []
        for label, (s, d) in all_data.items():
            def g소계(형태, _s=s):
                if _s.empty:
                    return {k:0 for k in ["청구금액","삭감액","심사결정금액","재심금액","불능보류금액","건수"]}
                row = _s[(_s["형태"]==형태) & (_s["보험자"]=="소계")]
                return row.iloc[0].to_dict() if not row.empty else {k:0 for k in ["청구금액","삭감액","심사결정금액","재심금액","불능보류금액","건수"]}
            외 = g소계("외래")
            입 = g소계("입원")
            tc = 외.get("청구금액",0) + 입.get("청구금액",0)
            td = 외.get("삭감액",0)   + 입.get("삭감액",0)
            rows_trend.append({
                "월": label,
                "총청구금액": tc, "총삭감액": td,
                "삭감률(%)": round(td/tc*100,2) if tc else 0,
                "외래건수": 외.get("건수",0), "입원건수": 입.get("건수",0),
                "재심금액": 외.get("재심금액",0)+입.get("재심금액",0),
                "불능보류금액": 외.get("불능보류금액",0)+입.get("불능보류금액",0),
            })
        trend = pd.DataFrame(rows_trend)
        fig_t = make_subplots(rows=2, cols=2,
                              subplot_titles=["총 청구금액 추이","삭감률(%) 추이",
                                             "외래/입원 건수","재심·불능보류 금액"])
        fig_t.add_trace(go.Scatter(x=trend["월"],y=trend["총청구금액"],mode="lines+markers",
                                   name="청구금액",line=dict(color="#2e6da4",width=2)),row=1,col=1)
        fig_t.add_trace(go.Scatter(x=trend["월"],y=trend["총삭감액"],mode="lines+markers",
                                   name="삭감액",line=dict(color="#e74c3c",width=2)),row=1,col=1)
        fig_t.add_trace(go.Scatter(x=trend["월"],y=trend["삭감률(%)"],
                                   mode="lines+markers+text",
                                   text=[f"{v:.2f}%" for v in trend["삭감률(%)"]],
                                   textposition="top center",name="삭감률",
                                   line=dict(color="#e67e22",width=2)),row=1,col=2)
        fig_t.add_hline(y=삭감률_임계값, line_dash="dash", line_color="red",
                        annotation_text=f"기준 {삭감률_임계값}%", row=1, col=2)
        fig_t.add_trace(go.Bar(x=trend["월"],y=trend["외래건수"],
                               name="외래건수",marker_color="#2e6da4"),row=2,col=1)
        fig_t.add_trace(go.Bar(x=trend["월"],y=trend["입원건수"],
                               name="입원건수",marker_color="#27ae60"),row=2,col=1)
        fig_t.add_trace(go.Bar(x=trend["월"],y=trend["재심금액"],
                               name="재심금액",marker_color="#8e44ad"),row=2,col=2)
        fig_t.add_trace(go.Bar(x=trend["월"],y=trend["불능보류금액"],
                               name="불능보류금액",marker_color="#e67e22"),row=2,col=2)
        fig_t.update_layout(height=620, barmode="group",
                            legend=dict(orientation="h",y=-0.1),
                            margin=dict(t=40,b=0))
        st.plotly_chart(fig_t, use_container_width=True)
        st.dataframe(trend, use_container_width=True, hide_index=True)


# ── 푸터 ──────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#b0b8c1;font-size:0.8rem'>"
    "🏥 청구심사 종합 분석 대시보드 | 제작자 : 주식회사 메디엄 조정윤 | Powered by Streamlit + Plotly"
    "</div>",
    unsafe_allow_html=True,
)
