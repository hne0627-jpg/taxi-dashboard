"""
개인택시 운행일지 대시보드
- 영수증(미터기/가스) 사진 올리면 자동 추출해서 구글 시트에 기록
- 순수익, 연비, 실차율 등 자동 계산 + 그래프 + 캘린더 조회 + 목표 게이지
- 폰 브라우저에서도 잘 보이게 세로 한 줄 구조
"""

import json
import base64
import calendar
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

# ----------------------------------------------------------------------------
# 기본 설정
# ----------------------------------------------------------------------------
st.set_page_config(page_title="운행일지", page_icon="🚕", layout="centered")

# 시트에 실제로 저장하는 칸 (원본 값만 저장, 계산값은 화면에서 만듦)
COLUMNS = [
    "날짜", "실차Km", "운행시간", "수입", "지출",
    "앱", "카드", "현금", "가스금액", "가스리터",
    "금일운행Km", "빈차Km", "누적Km",
]
NUMERIC = [c for c in COLUMNS if c != "날짜"]

DEFAULT_GOAL = 4_500_000  # 월 목표금액
CSV_FALLBACK = "migrated_taxi_2026.csv"

# 각 지표 도움말 (물음표 눌렀을 때 뜨는 설명)
HELP = {
    "순수익": "수입에서 가스비를 포함한 모든 지출을 뺀 금액입니다.\n\n"
            "하루에 실제로 손에 남은 돈입니다.",
    "시간당순수익": "한 시간 운행으로 실제 남긴 금액입니다.\n\n"
                "순수익을 운행 시간으로 나눠 계산합니다.",
    "누적순수익": "선택한 기간 동안 실제로 벌어들인 금액의 합계입니다.\n\n"
              "지출을 모두 뺀 금액입니다.",
    "리터당단가": "가스를 1리터에 얼마에 넣었는지 보여주는 금액입니다.",
    "연비": "가스 1리터로 달린 거리입니다.\n\n숫자가 클수록 연료 효율이 좋습니다.",
    "Km당연료비": "1km를 달리는 데 든 가스비입니다.\n\n숫자가 작을수록 효율이 좋습니다.",
    "Km당수입": "실차 주행거리 1km마다 벌어들인 금액입니다.",
    "시간당수입": "한 시간 운행할 때마다 벌어들인 금액입니다.",
    "실차율": "총 주행거리 중 실차 주행거리(손님 태운 거리)의 비율입니다.\n\n"
            "높을수록 빈 차로 다닌 시간이 적었다는 뜻입니다.",
    "공차율": "총 주행거리 중 공차 주행거리(빈 차로 간 거리)의 비율입니다.\n\n"
            "예전 시트의 '실차율' 칸이 사실 이 값이라, 이 숫자가 예전 기록과 같습니다.",
    "요일별수입": "요일별 평균 수입입니다.\n\n어떤 요일에 많이 버는지 확인할 수 있습니다.",
    "결제비율": "앱, 카드, 현금 결제 금액의 비율입니다.",
    "근무일": "선택한 기간 중 실제로 운행해 수입이 발생한 날의 수입니다.",
    "목표게이지": "이번 달 목표 금액 대비 지금까지의 수입 달성률입니다.",
}

WEEKDAY = ["월", "화", "수", "목", "금", "토", "일"]


def secret(key, default=None):
    """secrets.toml 이 없어도 안전하게 값 읽기 (없으면 기본값)."""
    try:
        return st.secrets[key]
    except Exception:
        return default


# ----------------------------------------------------------------------------
# 색과 그래프 공통 스타일 (색맹 안전 팔레트로 검증됨)
# ----------------------------------------------------------------------------
PRIMARY = "#2a78d6"                              # 기본 파랑 (단일 지표)
PAY_COLORS = ["#2a78d6", "#eb6834", "#1baf7a"]   # 앱, 카드, 현금
GRID = "#ecebe6"
AXIS = "#c3c2b7"
INK = "#52514e"
FONT = "Pretendard, system-ui, sans-serif"
# staticPlot: 그래프가 터치(확대/이동)를 가로채지 않게 정지 상태로. 페이지 스크롤이 자연스러워짐
CHART_CFG = {"displayModeBar": False, "staticPlot": True}


def style_fig(fig, height=240, day_axis=False):
    """그래프 공통 스타일: 여백, 격자, 축, 글꼴 일관되게."""
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=8, t=6, b=6),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color=INK, size=13),
        xaxis_title=None, yaxis_title=None,
        showlegend=False,
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, showline=True, linecolor=AXIS,
                     tickcolor=AXIS)
    fig.update_yaxes(showgrid=True, gridcolor=GRID, zeroline=False,
                     showline=False, tickformat=",d", rangemode="tozero")
    if day_axis:
        fig.update_xaxes(tickformat="%d", hoverformat="%m월 %d일",
                         nticks=12, ticklabelmode="period")
    return fig

# ----------------------------------------------------------------------------
# 구글 시트 연결 (연결 안 돼 있으면 CSV로 읽기 전용 동작)
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_worksheet():
    """구글 시트 워크시트 객체. secrets 없으면 None."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        info = dict(st.secrets["gcp_service_account"])
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["sheet_id"])
        ws = sh.sheet1
        # 헤더 없으면 한 번 세팅
        if ws.row_count == 0 or ws.acell("A1").value != "날짜":
            ws.update("A1", [COLUMNS])
        return ws
    except Exception:
        return None


def load_df():
    """시트(또는 CSV)에서 데이터를 DataFrame으로."""
    ws = get_worksheet()
    if ws is not None:
        records = ws.get_all_records()
        df = pd.DataFrame(records)
    else:
        try:
            df = pd.read_csv(CSV_FALLBACK)
        except Exception:
            df = pd.DataFrame(columns=COLUMNS)

    for c in COLUMNS:
        if c not in df.columns:
            df[c] = None
    df = df[COLUMNS].copy()
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    for c in NUMERIC:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["날짜"]).sort_values("날짜").reset_index(drop=True)
    return df


def save_row(row: dict):
    """한 날짜 기록을 시트에 추가하거나 갱신."""
    ws = get_worksheet()
    if ws is None:
        return False, "구글 시트가 연결되지 않아 저장할 수 없습니다. 현재는 확인만 가능합니다."
    values = ws.get_all_values()
    target_date = row["날짜"]
    existing, row_idx = None, None
    for i, r in enumerate(values[1:], start=2):
        if r and r[0] == target_date:
            existing = r
            row_idx = i
            break

    if row_idx:
        # 같은 날짜면 수정: 폼에 있는 칸은 새 값으로 덮어쓰고, 폼에 없는 칸은 기존 값 유지
        old = {COLUMNS[j]: (existing[j] if j < len(existing) else "")
               for j in range(len(COLUMNS))}
        line = [row[c] if c in row else old.get(c, "") for c in COLUMNS]
        ws.update(f"A{row_idx}", [line])
        return True, "수정되었습니다."
    else:
        line = [row.get(c, "") for c in COLUMNS]
        ws.append_row(line, value_input_option="USER_ENTERED")
        return True, "기록이 저장되었습니다."


# ----------------------------------------------------------------------------
# 파생 지표 계산
# ----------------------------------------------------------------------------
def add_metrics(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    if d.empty:
        return d

    def safe_div(a, b):
        return (a / b).where((b != 0) & b.notna())

    d["순수익"] = d["수입"].fillna(0) - d["지출"].fillna(0)
    d["시간당순수익"] = safe_div(d["순수익"], d["운행시간"])
    d["리터당단가"] = safe_div(d["지출"], d["가스리터"])
    d["연비"] = safe_div(d["금일운행Km"], d["가스리터"])
    d["Km당연료비"] = safe_div(d["지출"], d["금일운행Km"])
    d["Km당수입"] = safe_div(d["수입"], d["실차Km"])
    d["시간당수입"] = safe_div(d["수입"], d["운행시간"])
    d["실차율"] = safe_div(d["실차Km"], d["금일운행Km"]) * 100
    d["공차율"] = safe_div(d["빈차Km"], d["금일운행Km"]) * 100
    d["요일"] = d["날짜"].dt.weekday.map(lambda x: WEEKDAY[x])
    return d


# ----------------------------------------------------------------------------
# 영수증 추출 (Claude 비전)
# ----------------------------------------------------------------------------
EXTRACT_PROMPT = """이 이미지는 택시 영수증입니다. 두 종류 중 하나예요.
1) 미터기 일일 영수증: 운행거리, 운행시간, 수입, 결제수단(앱/카드/현금) 등
2) 가스(LPG) 충전 영수증: 충전 금액, 리터

이미지에서 읽을 수 있는 값만 뽑아서 아래 JSON 형식으로만 답하세요.
없는 값은 null로 두세요. 숫자는 콤마 없이 숫자만.

{
  "receipt_type": "meter 또는 gas",
  "날짜": "YYYY-MM-DD 또는 null",
  "실차Km": null,
  "운행시간": null,
  "수입": null,
  "지출": null,
  "앱": null,
  "카드": null,
  "현금": null,
  "가스금액": null,
  "가스리터": null,
  "금일운행Km": null,
  "빈차Km": null,
  "누적Km": null
}

실차Km는 영업거리, 금일운행Km는 총주행거리, 빈차Km는 공차거리입니다.
운행시간은 시간 단위 숫자(예: 3.25)로. JSON 외 다른 말은 하지 마세요."""


def extract_receipt(image_bytes: bytes, media_type: str) -> dict:
    import anthropic

    client = anthropic.Anthropic(api_key=st.secrets["anthropic_api_key"])
    model = secret("model", "claude-sonnet-5")
    b64 = base64.standard_b64encode(image_bytes).decode()
    msg = client.messages.create(
        model=model,
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": EXTRACT_PROMPT},
            ],
        }],
    )
    text = msg.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("```")[1].replace("json", "", 1).strip()
    return json.loads(text)


# ----------------------------------------------------------------------------
# 화면
# ----------------------------------------------------------------------------
# 글꼴, 여백, 큰 글씨 입력 편의, 넓은 화면 대응 스타일
st.markdown("""
<style>
  @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');
  html, body, [class*="css"], .stMarkdown, button, input, select {
    font-family: 'Pretendard', system-ui, -apple-system, sans-serif;
  }
  .block-container {padding: 1rem 1.4rem 3rem; max-width: 1100px;}
  #MainMenu, footer, header {visibility: hidden;}
  h1 {font-size: 1.6rem; font-weight: 800; letter-spacing: -0.02em;}
  h2 {font-size: 1.35rem !important; font-weight: 800;
    letter-spacing: -0.02em; margin: 0.4rem 0 0.2rem;}
  h3 {font-size: 1.18rem !important; font-weight: 700;
    letter-spacing: -0.01em; margin: 1rem 0 0.2rem;}

  /* 입력 라벨: 크고 진하게 */
  [data-testid="stWidgetLabel"] p, .stNumberInput label, .stDateInput label {
    font-size: 1.1rem !important; font-weight: 700 !important; color: #1f1f1f !important;
  }
  /* 입력칸: 글씨 크게, 칸 높게 (안경 없이도 잘 보이게) */
  .stNumberInput input, .stDateInput input, [data-baseweb="input"] input {
    font-size: 1.4rem !important; padding: 0.55rem 0.7rem !important;
    height: auto !important; font-weight: 600;
  }
  /* 숫자 증감 버튼 숨김 (직접 입력이 빠르고 화면도 깔끔) */
  [data-testid="stNumberInputStepUp"], [data-testid="stNumberInputStepDown"] {
    display: none !important;
  }
  /* 저장 등 버튼: 크고 누르기 쉽게 */
  .stButton button, .stFormSubmitButton button {
    font-size: 1.25rem !important; font-weight: 700 !important;
    padding: 0.75rem 1rem !important; border-radius: 12px !important;
  }
  /* 탭: 손가락으로 누르기 쉽게 */
  .stTabs [data-baseweb="tab-list"] {gap: 8px;}
  .stTabs [data-baseweb="tab"] {font-size: 1.05rem; font-weight: 700; padding: 0.4rem 0.2rem;}
  /* 접기 메뉴 제목도 크게 */
  .streamlit-expanderHeader, [data-testid="stExpander"] summary {
    font-size: 1.05rem !important; font-weight: 600 !important;
  }
  /* 미리보기 등 숫자 크게 */
  [data-testid="stMetricValue"] {font-size: 1.7rem; font-weight: 800;
    letter-spacing: -0.01em;}
  [data-testid="stMetricLabel"] p {color: #52514e; font-size: 0.95rem;}
  [data-testid="stMetricDelta"] {font-size: 0.85rem;}

  /* 지표를 카드로: 하얀 배경, 둥근 모서리, 옅은 테두리 */
  [data-testid="stMetric"] {
    background: #ffffff;
    border: 1px solid #e7e5df;
    border-radius: 16px;
    padding: 15px 18px;
    box-shadow: 0 1px 2px rgba(20, 20, 20, 0.04);
  }
  [data-testid="stMetricValue"] {font-variant-numeric: tabular-nums;}
  /* 그래프도 카드 안에 */
  [data-testid="stPlotlyChart"] {
    background: #ffffff;
    border: 1px solid #e7e5df;
    border-radius: 16px;
    padding: 10px 8px 4px;
    box-shadow: 0 1px 2px rgba(20, 20, 20, 0.04);
  }
  /* 입력칸도 하얗고 둥글게 */
  .stNumberInput input, .stDateInput input, [data-baseweb="input"] {
    background: #ffffff !important; border-radius: 12px !important;
  }
  [data-baseweb="input"] {border: 1px solid #e2e0da !important;}
  /* 탭 밑줄, 선택 강조를 파랑으로 */
  .stTabs [data-baseweb="tab-highlight"] {background-color: #2a78d6 !important;}

  /* 그래프가 두 개 나란한 줄은 좁은 화면에서 위아래로 쌓이게 */
  @media (max-width: 768px) {
    [data-testid="stHorizontalBlock"]:has(.stPlotlyChart) {flex-wrap: wrap;}
    [data-testid="stHorizontalBlock"]:has(.stPlotlyChart) > [data-testid="column"] {
      min-width: 100% !important; flex: 1 1 100% !important;
    }
  }
</style>
""", unsafe_allow_html=True)

st.title("🚕 운행일지")

df_all = add_metrics(load_df())
sheet_ready = get_worksheet() is not None
if not sheet_ready:
    st.info("구글 시트 연결 전에는 지난 기록만 확인할 수 있습니다. "
            "연결을 완료하면 기록을 바로 저장할 수 있습니다.")

goal = int(secret("monthly_goal", DEFAULT_GOAL))

tab_input, tab_dash, tab_log = st.tabs(["✏️ 기록하기", "📊 대시보드", "📅 지난 기록"])

# ---- 탭 1: 기록 입력 (수동) -------------------------------------------------
with tab_input:
    st.header("오늘 운행 기록", anchor=False,
              help="기존에 작성하시던 항목만 입력하시면 됩니다.\n\n"
                   "실차율과 주행거리는 자동으로 계산됩니다.")
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))

    # 이번 달 진행 상황 (기록하러 들어오면 한 줄로 한눈에)
    _tm = date.today().strftime("%Y-%m")
    _cur = df_all[df_all["날짜"].dt.strftime("%Y-%m") == _tm] if not df_all.empty else df_all
    _income = int(_cur["수입"].fillna(0).sum()) if not _cur.empty else 0
    _days = int((_cur["수입"].fillna(0) > 0).sum()) if not _cur.empty else 0
    _pct = round(_income / goal * 100) if goal else 0
    st.markdown(
        f"<div style='background:#eef4fd;border:1px solid #d8e6fb;border-radius:12px;"
        f"padding:10px 14px;margin:2px 0 6px;font-size:1rem;color:#1b1b1a;'>"
        f"이번 달 <b>{_income:,}원</b> 벌었어요. 목표의 <b>{_pct}%</b>, {_days}일 근무했어요."
        f"</div>", unsafe_allow_html=True)

    # 날짜는 폼 밖에 둠: 날짜를 바꾸면 그날 저장된 값을 불러와 칸에 채움
    in_date = st.date_input("운행 날짜", value=date.today())

    existing = None
    if not df_all.empty:
        _match = df_all[df_all["날짜"].dt.date == in_date]
        if not _match.empty:
            existing = _match.iloc[0]
    editing = existing is not None

    def _v(col, default=0.0):
        if existing is not None and pd.notna(existing.get(col)):
            return float(existing.get(col))
        return default

    _hh = int(_v("운행시간"))
    _mm = int(round((_v("운행시간") - _hh) * 60))
    if _mm >= 60:  # 반올림으로 60이 되면 시간으로 올림
        _hh += 1
        _mm -= 60

    if editing:
        st.info("이 날짜에 이미 기록이 있어요. 아래 값을 고쳐서 저장하면 수정됩니다.")

    # 폼 이름에 날짜를 넣어, 날짜가 바뀌면 폼이 새로 그려지며 그날 값으로 채워짐
    # 원본 파일 순서대로 손입력 칸만 배치 (공차/누적/실차율은 자동이라 표에서 확인)
    with st.form(f"day_form_{in_date.isoformat()}", clear_on_submit=False):
        실차Km = st.number_input("실차 주행거리 (km)", min_value=0.0, step=1.0, value=_v("실차Km"),
                                help="손님을 태우고 달린 거리예요.")

        st.subheader("운행 시간", anchor=False)
        t1, t2 = st.columns(2)
        운행시 = t1.number_input("시간", min_value=0, step=1, value=_hh)
        운행분 = t2.number_input("분", min_value=0, max_value=59, step=5, value=_mm)

        수입 = st.number_input("수입 (원)", min_value=0, step=1000, value=int(_v("수입")))
        지출 = st.number_input("지출 (가스 금액, 원)", min_value=0, step=1000, value=int(_v("지출")),
                              help="주유한 날 가스 영수증 금액을 넣으세요.\n\n이 금액이 그날 지출로 반영됩니다.")
        가스리터 = st.number_input("충전량 (L)", min_value=0.0, step=1.0, value=_v("가스리터"),
                                help="주유한 날 넣은 가스 양이에요. 연비 계산에 쓰여요.")

        st.subheader("결제수단별 수입", anchor=False)
        p1, p2, p3 = st.columns(3)
        앱 = p1.number_input("앱 (원)", min_value=0, step=1000, value=int(_v("앱")))
        카드 = p2.number_input("카드 (원)", min_value=0, step=1000, value=int(_v("카드")))
        현금 = p3.number_input("현금 (원)", min_value=0, step=1000, value=int(_v("현금")))

        금일Km = st.number_input("총 주행거리 (km)", min_value=0.0, step=1.0, value=_v("금일운행Km"),
                                help="오늘 실제로 달린 전체 거리예요 (실차 + 공차).\n\n"
                                     "공차 주행거리는 여기서 실차를 빼서 자동 계산됩니다.")

        st.caption("공차 주행거리, 누적 주행거리, 실차율은 저장하면 자동으로 계산돼 "
                   "'지난 기록'에서 볼 수 있어요.")
        _btn = "수정 저장하기" if editing else "기록 저장하기"
        submitted = st.form_submit_button(_btn, type="primary", use_container_width=True)

    if submitted:
        운행시간 = round(운행시 + 운행분 / 60, 4)
        빈차Km = max(금일Km - 실차Km, 0.0)
        if editing and pd.notna(existing.get("누적Km")):
            누적Km = int(existing.get("누적Km"))  # 수정 시 누적거리는 그대로 유지
        else:
            last_series = df_all["누적Km"].dropna() if not df_all.empty else pd.Series(dtype=float)
            last_nujeok = int(last_series.iloc[-1]) if len(last_series) else 0
            누적Km = last_nujeok + int(금일Km) if 금일Km else ""
        순수익 = int(수입) - int(지출)
        if 수입 == 0 and 금일Km == 0 and 지출 == 0:
            st.warning("입력한 값이 없습니다. 수입을 입력한 후 저장해주세요.")
        else:
            row = {
                "날짜": in_date.isoformat(),
                "실차Km": 실차Km or "", "운행시간": 운행시간 or "",
                "수입": 수입 or "", "지출": 지출 or "",
                "앱": 앱 or "", "카드": 카드 or "", "현금": 현금 or "",
                "가스리터": 가스리터 or "",
                "금일운행Km": 금일Km or "", "빈차Km": 빈차Km or "",
                "누적Km": 누적Km or "",
            }
            ok, msg = save_row(row)
            if ok:
                verb = "수정" if editing else "저장"
                st.session_state["flash"] = f"{verb}되었습니다. 순수익 {순수익:,}원."
                st.rerun()
            else:
                st.error(msg)

# ---- 탭 2: 대시보드 ---------------------------------------------------------
with tab_dash:
    if df_all.empty:
        st.warning("표시할 데이터가 아직 없습니다.")
    else:
        months = sorted(df_all["날짜"].dt.strftime("%Y-%m").unique(), reverse=True)
        labels = {mk: f"{int(mk[:4])}년 {int(mk[5:7])}월" for mk in months}
        sel = st.selectbox("월 선택", months, index=0,
                           format_func=lambda mk: labels[mk],
                           label_visibility="collapsed")
        m = df_all[df_all["날짜"].dt.strftime("%Y-%m") == sel]

        # 지난달 데이터 (증감 표시용)
        idx = months.index(sel)
        prev = (df_all[df_all["날짜"].dt.strftime("%Y-%m") == months[idx + 1]]
                if idx + 1 < len(months) else None)

        earned = int(m["수입"].fillna(0).sum())
        net = int(m["순수익"].sum())
        work_days = int((m["수입"].fillna(0) > 0).sum())
        pct = round(earned / goal * 100) if goal else 0

        # 목표 게이지 (월은 위 선택칸으로 끝, 여기선 반복 안 함)
        st.subheader(f"이번 달 목표 달성률 {pct}% (목표 {goal:,}원)",
                     help=HELP["목표게이지"], anchor=False)
        gfig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=earned,
            number={"prefix": "₩", "valueformat": ",",
                    "font": {"size": 30, "color": INK}},
            gauge={
                "axis": {"range": [0, goal], "tickformat": ".2s",
                         "tickcolor": AXIS, "tickfont": {"size": 10, "color": AXIS}},
                "bar": {"color": PRIMARY, "thickness": 0.72},
                "bgcolor": GRID, "borderwidth": 0,
                "threshold": {"line": {"color": "#d03b3b", "width": 3},
                              "value": goal},
            },
        ))
        gfig.update_layout(height=190, margin=dict(l=24, r=24, t=6, b=0),
                           font=dict(family=FONT))
        st.plotly_chart(gfig, use_container_width=True, config=CHART_CFG)

        # 목표 페이스 안내 (이번 달이면 남은 기간 하루 필요액까지)
        today = date.today()
        remain = goal - earned
        sel_y, sel_m = int(sel[:4]), int(sel[5:7])
        if remain <= 0:
            st.success(f"목표를 {abs(remain):,}원 넘겼어요. 잘하고 계세요!")
        elif (sel_y, sel_m) == (today.year, today.month):
            days_left = calendar.monthrange(sel_y, sel_m)[1] - today.day + 1
            per_day = remain / days_left if days_left else remain
            st.info(f"목표까지 {remain:,}원 남았어요. "
                    f"남은 {days_left}일 동안 하루 {per_day:,.0f}원이면 달성해요.")
        else:
            st.caption(f"이 달은 목표보다 {remain:,}원 적었어요.")

        # 이번 달 요약
        def delta(cur, sumfn):
            if prev is None or prev.empty:
                return None
            p = sumfn(prev)
            return f"{cur - p:+,}" if p else None

        c1, c2, c3 = st.columns(3)
        c1.metric("총 수입", f"{earned:,}원",
                  delta=delta(earned, lambda d: int(d["수입"].fillna(0).sum())),
                  help="이번 달 수입의 합계입니다. 아래 숫자는 지난달과 비교한 증감입니다.")
        c2.metric("순수익", f"{net:,}원",
                  delta=delta(net, lambda d: int(d["순수익"].sum())),
                  help=HELP["누적순수익"])
        c3.metric("근무일", f"{work_days}일", help=HELP["근무일"])

        st.subheader("효율 지표", anchor=False)

        def avg(col):
            s = m[col].replace([float("inf"), float("-inf")], pd.NA).dropna()
            return s.mean() if len(s) else None

        def fmt(v, dec=0, suffix=""):
            return "-" if v is None else f"{v:,.{dec}f}{suffix}"

        e1, e2, e3 = st.columns(3)
        e1.metric("시간당 순수익", fmt(avg("시간당순수익"), 0, "원"), help=HELP["시간당순수익"])
        e2.metric("시간당 수입", fmt(avg("시간당수입"), 0, "원"), help=HELP["시간당수입"])
        e3.metric("1km당 수입", fmt(avg("Km당수입"), 0, "원"), help=HELP["Km당수입"])
        e4, e5 = st.columns(2)
        e4.metric("실차율", fmt(avg("실차율"), 0, "%"), help=HELP["실차율"])
        e5.metric("공차율", fmt(avg("공차율"), 0, "%"), help=HELP["공차율"])
        e6, e7 = st.columns(2)
        e6.metric("연비", fmt(avg("연비"), 1, " km/L"), help=HELP["연비"])
        e7.metric("리터당 가스비", fmt(avg("리터당단가"), 0, "원"), help=HELP["리터당단가"])

        st.divider()

        # 일별 순수익 (전체 너비). 운행 안 한 날은 회색 점으로 표시
        st.subheader("날짜별 순수익", anchor=False,
                     help="날짜별 순수익입니다.\n\n막대가 높을수록 그날 많이 남았다는 뜻입니다.")
        bar = px.bar(m, x="날짜", y="순수익", color_discrete_sequence=[PRIMARY])
        bar.update_traces(marker=dict(cornerradius=5),
                          hovertemplate="순수익 %{y:,}원<extra></extra>")
        rest = m[m["수입"].fillna(0) == 0]
        if not rest.empty:
            bar.add_scatter(x=rest["날짜"], y=[0] * len(rest), mode="markers",
                            marker=dict(color="#b9b8b1", size=7),
                            hovertemplate="쉬는 날<extra></extra>", showlegend=False)
        style_fig(bar, 300, day_axis=True)
        st.plotly_chart(bar, use_container_width=True, config=CHART_CFG)
        st.caption("막대가 없고 회색 점만 있는 날은 운행하지 않은 쉬는 날입니다.")

        # 실차율 추이 (전체 너비). 운행한 날만 이어서 표시
        st.subheader("실차율 추이", help=HELP["실차율"], anchor=False)
        line = px.line(m, x="날짜", y="실차율", markers=True,
                       color_discrete_sequence=[PRIMARY])
        line.update_traces(line_width=2, marker_size=6, connectgaps=True,
                           hovertemplate="실차율 %{y:.0f}%<extra></extra>")
        style_fig(line, 280, day_axis=True)
        line.update_yaxes(ticksuffix="%")
        st.plotly_chart(line, use_container_width=True, config=CHART_CFG)
        st.caption("운행한 날만 이어서 표시됩니다. 점이 없는 날은 쉬는 날입니다.")

        # 요일별 평균 수입 + 결제수단 비율 (넓은 화면에선 나란히, 좁으면 위아래)
        left, right = st.columns(2)
        with left:
            st.subheader("요일별 평균 수입", help=HELP["요일별수입"], anchor=False)
            wk = (m[m["수입"].fillna(0) > 0].groupby("요일")["수입"].mean()
                  .reindex(WEEKDAY).reset_index())
            wk["수입"] = wk["수입"].fillna(0)
            wbar = px.bar(wk, x="요일", y="수입", color_discrete_sequence=[PRIMARY])
            wbar.update_traces(marker=dict(cornerradius=5),
                               hovertemplate="%{x}요일 평균 %{y:,.0f}원<extra></extra>")
            style_fig(wbar, 300)
            st.plotly_chart(wbar, use_container_width=True, config=CHART_CFG)
            st.caption("막대가 없는 요일은 이 달에 운행하지 않은 요일입니다.")
        with right:
            st.subheader("결제수단 비율", help=HELP["결제비율"], anchor=False)
            pay = pd.DataFrame({
                "수단": ["앱", "카드", "현금"],
                "금액": [m["앱"].fillna(0).sum(), m["카드"].fillna(0).sum(),
                        m["현금"].fillna(0).sum()],
            })
            pie = px.pie(pay, names="수단", values="금액", hole=0.55, color="수단",
                         color_discrete_map={"앱": PAY_COLORS[0], "카드": PAY_COLORS[1],
                                             "현금": PAY_COLORS[2]})
            pie.update_traces(textinfo="label+percent", textposition="outside",
                              marker=dict(line=dict(color="#ffffff", width=2)),
                              hovertemplate="%{label} %{value:,}원 (%{percent})<extra></extra>")
            pie.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10),
                              showlegend=False, font=dict(family=FONT, color=INK))
            st.plotly_chart(pie, use_container_width=True, config=CHART_CFG)

# ---- 탭 3: 기록 조회 --------------------------------------------------------
with tab_log:
    if df_all.empty:
        st.warning("표시할 데이터가 아직 없습니다.")
    else:
        st.subheader("기간별 기록")
        st.caption("아래 버튼으로 빠르게 선택하거나, 달력에서 기간을 직접 선택할 수 있습니다.")
        dmin = df_all["날짜"].min().date()
        dmax = df_all["날짜"].max().date()

        # 빠른 선택 버튼
        preset = st.segmented_control(
            "빠른 선택",
            ["최근 7일", "최근 30일", "이번 달", "지난달", "올해", "전체"],
            default="이번 달", label_visibility="collapsed") or "이번 달"

        if preset == "최근 7일":
            s, e = max(dmin, dmax - timedelta(days=6)), dmax
        elif preset == "최근 30일":
            s, e = max(dmin, dmax - timedelta(days=29)), dmax
        elif preset == "지난달":
            last_prev = dmax.replace(day=1) - timedelta(days=1)
            s, e = last_prev.replace(day=1), last_prev
        elif preset == "올해":
            s, e = max(dmin, date(dmax.year, 1, 1)), dmax
        elif preset == "전체":
            s, e = dmin, dmax
        else:  # 이번 달
            s, e = max(dmin, dmax.replace(day=1)), dmax

        # 달력 하나에서 시작일과 종료일을 함께 선택 (버튼을 바꾸면 따라 바뀜)
        rng = st.date_input("조회 기간", value=(s, e), min_value=dmin, max_value=dmax,
                            key=f"range_{preset}", format="YYYY.MM.DD")
        if isinstance(rng, (tuple, list)) and len(rng) == 2:
            s, e = rng
        elif isinstance(rng, (tuple, list)) and len(rng) == 1:
            s = e = rng[0]

        f = df_all[(df_all["날짜"].dt.date >= s) & (df_all["날짜"].dt.date <= e)]
        k1, k2, k3 = st.columns(3)
        k1.metric("수입 합계", f"{int(f['수입'].fillna(0).sum()):,}원")
        k2.metric("순수익 합계", f"{int(f['순수익'].sum()):,}원", help=HELP["순수익"])
        k3.metric("근무일", f"{int((f['수입'].fillna(0) > 0).sum())}일", help=HELP["근무일"])

        # 표: 단위 붙여서 보기 좋게 (돈은 콤마, 퍼센트는 정수, 시간은 N시간 M분)
        def won(v):
            return f"{int(v):,}" if pd.notna(v) and v != "" else ""

        def pct(v):
            return f"{v:.0f}%" if pd.notna(v) else ""

        def hm(v):
            if pd.isna(v) or v == "":
                return ""
            h = int(v)
            mm = round((v - h) * 60)
            return f"{h}시간 {mm}분" if mm else f"{h}시간"

        def km(v):
            return f"{v:,.0f}" if pd.notna(v) and v != "" else ""

        def lit(v):
            return f"{v:.1f}" if pd.notna(v) and v != "" else ""

        def W(v):  # 돈 (없으면 -)
            s = won(v)
            return s + "원" if s else "-"

        def K(v):  # 거리 (없으면 -)
            s = km(v)
            return s + "km" if s else "-"

        def cell(lab, val):
            return (f"<div style='font-size:0.95rem;color:#6b6a66;padding:1px 0'>{lab} "
                    f"<b style='color:#1b1b1a'>{val}</b></div>")

        # 카드 뷰: 하루씩 위아래로 (모바일에서 옆으로 안 밀어도 됨). 최근 날짜가 위
        st.caption("최근 날짜가 위에 있어요. 운행한 날은 카드로 자세히, 쉬는 날은 한 줄로 보여요.")
        cards = []
        for _, r in f.sort_values("날짜", ascending=False).iterrows():
            dstr = f"{r['날짜'].month}월 {r['날짜'].day}일 ({r['요일']})"
            inc = float(r["수입"]) if pd.notna(r["수입"]) else 0.0
            if inc <= 0:
                cards.append(
                    "<div style='border:1px solid #ebe9e3;border-radius:12px;background:#fbfbf9;"
                    "padding:9px 14px;margin-bottom:8px'>"
                    f"<b style='color:#8a8880'>{dstr}</b>"
                    "<span style='color:#a8a69e;margin-left:8px'>운행 안 함</span></div>")
                continue
            net = int(r["순수익"]) if pd.notna(r["순수익"]) else 0
            nc = "#2a78d6" if net >= 0 else "#d03b3b"
            body = "".join([
                cell("수입", W(r["수입"])), cell("지출", W(r["지출"])),
                cell("운행", hm(r["운행시간"]) or "-"),
                cell("실차", K(r["실차Km"])), cell("공차", K(r["빈차Km"])),
                cell("총주행", K(r["금일운행Km"])),
                cell("실차율", pct(r["실차율"]) or "-"), cell("공차율", pct(r["공차율"]) or "-"),
                cell("충전", (lit(r["가스리터"]) + "L") if lit(r["가스리터"]) else "-"),
                cell("앱", W(r["앱"])), cell("카드", W(r["카드"])), cell("현금", W(r["현금"])),
                cell("누적", K(r["누적Km"])),
            ])
            cards.append(
                "<div style='border:1px solid #e7e5df;border-radius:14px;background:#fff;"
                "padding:12px 14px;margin-bottom:10px;box-shadow:0 1px 2px rgba(20,20,20,0.04)'>"
                "<div style='display:flex;justify-content:space-between;align-items:baseline;"
                "margin-bottom:6px'>"
                f"<span style='font-size:1.08rem;font-weight:800;color:#1b1b1a'>{dstr}</span>"
                f"<span style='font-size:1rem;font-weight:700;color:{nc}'>순수익 {net:,}원</span></div>"
                f"<div style='display:grid;grid-template-columns:1fr 1fr;column-gap:14px'>{body}</div>"
                "</div>")
        st.markdown(
            "".join(cards) or "<div style='color:#8a8880'>이 기간에 기록이 없어요.</div>",
            unsafe_allow_html=True)
