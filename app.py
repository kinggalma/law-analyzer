"""
건설안전기술사 법 분석 프로그램
- 법규 Sumary.pdf + 25년 이후 법개정 파일들을 기반으로 법령 조문을 검색한다.
- 실행: streamlit run app.py
- 데이터 준비: python extract_law.py 실행 후 data/law_data.xlsx 생성 필요
"""

import os
import re

import numpy as np
import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
import streamlit.components.v1 as components
import yaml
from yaml.loader import SafeLoader
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import plotly.express as px

# ── 상수 ─────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data", "law_data.xlsx")
CONFIG_PATH = os.path.join(BASE_DIR, "config.yaml")
SHEET_NAME = "법령조문"
REQUIRED_COLUMNS = [
    "법령명", "법령코드", "편장절", "조문번호", "조문제목",
    "조문내용", "시행일", "개정여부", "검색텍스트",
]
OPTIONAL_COLUMNS = ["개정전조문내용", "변경유형", "개정내용요약"]
EDITABLE_COLUMNS = [
    "법령명", "법령코드", "시행일", "개정여부", "변경유형", "개정내용요약",
]

LAW_COLORS = {
    "산안법":    "#1565C0",
    "산안기준":  "#2E7D32",
    "건진법":    "#E65100",
    "시설물법":  "#6A1B9A",
    "중처법":    "#B71C1C",
    "산안관리비": "#00695C",
    "콘크리트":  "#4E342E",
    "기타":     "#546E7A",
}

LAW_NAMES = {
    "산안법":    "산업안전보건법",
    "산안기준":  "산업안전보건기준에 관한 규칙",
    "건진법":    "건설기술진흥법",
    "시설물법":  "시설물의 안전 및 유지관리에 관한 특별법",
    "중처법":    "중대재해처벌법",
    "산안관리비": "산업안전보건관리비",
    "콘크리트":  "콘크리트공사 표준안전 작업지침",
    "기타":     "기타 법령",
}

AMENDMENT_BADGE = {
    "original": ("", "#9E9E9E"),
    "amended":  ("개정", "#FF6F00"),
    "new":      ("신설", "#C62828"),
}

LAW_INFERENCE_RULES = [
    ("산업안전보건기준에 관한 규칙", "산업안전보건기준에 관한 규칙", "산안기준"),
    ("산업안전보건기준", "산업안전보건기준에 관한 규칙", "산안기준"),
    ("산업안전보건법", "산업안전보건법", "산안법"),
    ("건설기술진흥법", "건설기술진흥법", "건진법"),
    ("시설물의 안전 및 유지관리에 관한 특별법", "시설물의 안전 및 유지관리에 관한 특별법", "시설물법"),
    ("시설물안전법", "시설물의 안전 및 유지관리에 관한 특별법", "시설물법"),
    ("중대재해 처벌 등에 관한 법률", "중대재해처벌법", "중처법"),
    ("중대재해처벌법", "중대재해처벌법", "중처법"),
    ("산업안전보건관리비", "산업안전보건관리비", "산안관리비"),
    ("콘크리트공사 표준안전 작업지침", "콘크리트공사 표준안전 작업지침", "콘크리트"),
    ("KCS 14 20 40", "콘크리트공사 표준안전 작업지침", "콘크리트"),
    ("한중콘크리트", "콘크리트공사 표준안전 작업지침", "콘크리트"),
    ("콘크리트", "콘크리트공사 표준안전 작업지침", "콘크리트"),
]

DATE_PATTERNS = [
    re.compile(r"(\d{4})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})"),
    re.compile(r"(\d{2})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})"),
]


# ── 인증 설정 로딩 ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def get_auth_config():
    """Streamlit Cloud secrets → 로컬 config.yaml 순으로 인증 설정 로드."""
    try:
        if "credentials" in st.secrets and "cookie" in st.secrets:
            credentials = {"usernames": {}}
            for uname, info in st.secrets["credentials"]["usernames"].items():
                credentials["usernames"][uname] = {
                    "name": info["name"],
                    "email": info.get("email", ""),
                    "password": info["password"],
                }
            return {
                "credentials": credentials,
                "cookie": {
                    "name": st.secrets["cookie"]["name"],
                    "key": st.secrets["cookie"]["key"],
                    "expiry_days": int(st.secrets["cookie"]["expiry_days"]),
                },
            }
    except Exception as exc:
        print(f"[auth] Streamlit secrets 로드 실패: {exc}")

    env_config = os.environ.get("AUTH_CONFIG_YAML")
    if env_config:
        try:
            return yaml.safe_load(env_config.replace("\\n", "\n"))
        except yaml.YAMLError as exc:
            print(f"[auth] AUTH_CONFIG_YAML 파싱 실패: {exc}")

    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding="utf-8-sig") as f:
            return yaml.load(f, Loader=SafeLoader)

    return None


st.set_page_config(
    page_title="건설안전기술사 법령 분석",
    page_icon="⚖️",
    layout="wide",
)


def prevent_browser_translation():
    """브라우저/번역 확장 기능이 전문 법령 용어를 오역하지 않도록 차단한다."""
    components.html(
        """
        <script>
        const doc = window.parent.document;
        const markNotranslate = () => {
            doc.documentElement.lang = "ko";
            doc.documentElement.setAttribute("translate", "no");
            doc.documentElement.classList.add("notranslate");
            doc.body.classList.add("notranslate");
            doc.body.setAttribute("translate", "no");
            doc.querySelectorAll("[data-testid='stAppViewContainer'], [data-testid='stSidebar']")
                .forEach((el) => {
                    el.classList.add("notranslate");
                    el.setAttribute("translate", "no");
                });
        };

        if (!doc.querySelector('meta[name="google"][content="notranslate"]')) {
            const meta = doc.createElement("meta");
            meta.name = "google";
            meta.content = "notranslate";
            doc.head.appendChild(meta);
        }

        markNotranslate();
        const observer = new MutationObserver(markNotranslate);
        observer.observe(doc.body, {childList: true, subtree: true});
        </script>
        """,
        height=0,
    )


# ── 데이터 로드 ───────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="법령 데이터 로딩 중...")
def load_data(_mtime: float | None = None) -> pd.DataFrame:
    if not os.path.exists(DATA_PATH):
        return pd.DataFrame()

    df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        st.error(f"데이터 파일에 필요한 컬럼이 없습니다: {', '.join(missing)}")
        return pd.DataFrame()

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    if "id" not in df.columns:
        df["id"] = range(1, len(df) + 1)

    df = df[df["조문내용"].ne("")]
    df["검색텍스트"] = (
        df["법령명"].fillna("") + " " +
        df["조문번호"].fillna("") + " " +
        df["조문제목"].fillna("") + " " +
        df["조문내용"].fillna("")
    )
    df = df.reset_index(drop=True)
    return df


def make_search_text(row: pd.Series) -> str:
    return " ".join(
        str(row.get(col, "")).strip()
        for col in ["법령명", "조문번호", "조문제목", "조문내용"]
        if str(row.get(col, "")).strip()
    )


def infer_law_info(row: pd.Series) -> tuple[str, str]:
    haystack = " ".join(
        str(row.get(col, ""))
        for col in ["법령명", "조문제목", "조문내용", "출처파일", "개정내용요약"]
    )
    for keyword, law_name, law_code in LAW_INFERENCE_RULES:
        if keyword in haystack:
            return law_name, law_code
    return str(row.get("법령명", "")).strip(), str(row.get("법령코드", "")).strip()


def infer_enforcement_date(row: pd.Series) -> str:
    haystack = " ".join(
        str(row.get(col, ""))
        for col in ["출처파일", "개정내용요약", "조문내용"]
    )
    for pat in DATE_PATTERNS:
        match = pat.search(haystack)
        if not match:
            continue
        year, month, day = match.groups()
        if len(year) == 2:
            year = f"20{year}"
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    return str(row.get("시행일", "")).strip()


def normalize_amendment(row: pd.Series) -> str:
    current = str(row.get("개정여부", "")).strip()
    if current in {"original", "amended", "new"}:
        return current
    if is_new_article(row):
        return "new"
    source = str(row.get("출처파일", "")).strip()
    if source and source != "법규 Sumary.pdf":
        return "amended"
    return "original"


def save_law_data(updated_df: pd.DataFrame) -> None:
    sheets = pd.read_excel(DATA_PATH, sheet_name=None)
    history_df = sheets.get("개정이력", pd.DataFrame())
    with pd.ExcelWriter(DATA_PATH, engine="openpyxl") as writer:
        updated_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        history_df.to_excel(writer, sheet_name="개정이력", index=False)


@st.cache_resource(show_spinner="TF-IDF 인덱스 구축 중...")
def build_tfidf(df: pd.DataFrame):
    texts = df["검색텍스트"].tolist()
    vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 3), min_df=1)
    matrix = vec.fit_transform(texts)
    return vec, matrix


# ── 검색 함수 ─────────────────────────────────────────────────────────────────

def search_keyword(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    mask = df["검색텍스트"].str.contains(keyword, case=False, na=False, regex=False)
    return df[mask].copy()


def search_similar(
    df: pd.DataFrame,
    vec: TfidfVectorizer,
    matrix,
    query: str,
    top_n: int = 50,
    threshold: float = 0.10,
) -> list[tuple[int, float]]:
    q_vec = vec.transform([query])
    sims = cosine_similarity(q_vec, matrix).flatten()
    top_idx = np.argsort(sims)[::-1]
    results = []
    for i in top_idx:
        if sims[i] >= threshold:
            results.append((int(i), round(float(sims[i]) * 100, 1)))
        if len(results) >= top_n:
            break
    return results


# ── UI 헬퍼 ──────────────────────────────────────────────────────────────────

def highlight(text: str, keyword: str) -> str:
    if not keyword or not text:
        return text
    escaped = re.escape(keyword)
    return re.sub(
        f"({escaped})",
        r'<mark style="background:#FFF176;padding:0 2px;border-radius:3px">\1</mark>',
        str(text),
        flags=re.IGNORECASE,
    )


def is_new_article(row: pd.Series) -> bool:
    status = str(row.get("개정여부", "")).strip()
    change_type = str(row.get("변경유형", "")).strip()
    summary = str(row.get("개정내용요약", "")).strip()
    return status == "new" or "신설" in change_type or "신설" in summary or "신규" in summary


def build_original_lookup(df: pd.DataFrame) -> dict[tuple[str, str], pd.Series]:
    original_df = df[df["개정여부"].eq("original")].copy()
    lookup = {}
    for _, item in original_df.iterrows():
        key = (str(item.get("법령명", "")).strip(), str(item.get("조문번호", "")).strip())
        if key[0] and key[1] and key not in lookup:
            lookup[key] = item
    return lookup


def render_comparison(row: pd.Series, keyword: str, original_lookup: dict[tuple[str, str], pd.Series] | None):
    amendment = str(row.get("개정여부", "original")).strip()
    if amendment == "original":
        return False

    content = str(row.get("조문내용", "")).strip()
    old_content = str(row.get("개정전조문내용", "")).strip()
    key = (str(row.get("법령명", "")).strip(), str(row.get("조문번호", "")).strip())

    if not old_content and original_lookup:
        original_row = original_lookup.get(key)
        if original_row is not None:
            old_content = str(original_row.get("조문내용", "")).strip()

    if is_new_article(row):
        st.success("신규로 추가된 조문입니다.")
        st.markdown(
            f'<div style="line-height:1.8;white-space:pre-wrap">{highlight(content, keyword)}</div>',
            unsafe_allow_html=True,
        )
        return True

    if not old_content:
        st.info("기존 조문 원문이 연결되지 않아 변경 조문만 표시합니다.")
        st.markdown(
            f'<div style="line-height:1.8;white-space:pre-wrap">{highlight(content, keyword)}</div>',
            unsafe_allow_html=True,
        )
        return True

    st.markdown("#### 기존 조문 / 변경 조문 비교")
    before_col, after_col = st.columns(2)
    with before_col:
        st.markdown("**기존 조문**")
        st.markdown(
            f'<div style="line-height:1.8;white-space:pre-wrap;background:#FAFAFA;'
            f'border:1px solid #E0E0E0;border-radius:6px;padding:12px">{highlight(old_content, keyword)}</div>',
            unsafe_allow_html=True,
        )
    with after_col:
        st.markdown("**변경 조문**")
        st.markdown(
            f'<div style="line-height:1.8;white-space:pre-wrap;background:#FFF8E1;'
            f'border:1px solid #FFE082;border-radius:6px;padding:12px">{highlight(content, keyword)}</div>',
            unsafe_allow_html=True,
        )
    return True


def render_law_card(
    row: pd.Series,
    keyword: str = "",
    score: float | None = None,
    original_lookup: dict[tuple[str, str], pd.Series] | None = None,
):
    law_code = row.get("법령코드", "기타")
    color = LAW_COLORS.get(law_code, "#546E7A")
    amendment = row.get("개정여부", "original")
    badge_text, badge_color = AMENDMENT_BADGE.get(amendment, ("", "#9E9E9E"))
    if is_new_article(row):
        badge_text, badge_color = "신규 추가", "#C62828"
    enforcement = row.get("시행일", "")

    # 헤더 뱃지 라인
    law_badge = (
        f'<span style="background:{color};color:#fff;padding:2px 8px;'
        f'border-radius:4px;font-size:0.8em">{row.get("법령명", law_code)}</span>'
    )
    amend_badge = (
        f'<span style="background:{badge_color};color:#fff;padding:2px 8px;'
        f'border-radius:4px;font-size:0.8em;margin-left:6px">{badge_text}</span>'
        if badge_text else ""
    )
    score_badge = (
        f'<span style="background:#eee;padding:2px 8px;border-radius:4px;'
        f'font-size:0.8em;margin-left:6px">유사도 {score}%</span>'
        if score is not None else ""
    )
    date_badge = (
        f'<span style="background:#F5F5F5;color:#555;padding:2px 8px;'
        f'border-radius:4px;font-size:0.8em;margin-left:6px">시행 {enforcement}</span>'
        if enforcement and enforcement != "미상" else ""
    )

    art_no = str(row.get("조문번호", "")).strip()
    art_title = str(row.get("조문제목", "")).strip()
    content = str(row.get("조문내용", "")).strip()
    chungjang = str(row.get("편장절", "")).strip()
    law_name_label = str(row.get("법령명", law_code)).strip()

    # 법령명 + 조문번호 + 조문제목 순으로 표시: "산업안전보건법  제41조  안전조치"
    label_parts = [law_name_label, art_no, art_title or content[:30]]
    label = "  ".join(filter(None, label_parts))
    if score is not None:
        label += f"  (유사도 {score}%)"

    with st.expander(f"⚖️ {label}"):
        st.markdown(
            f"{law_badge}{amend_badge}{score_badge}{date_badge}",
            unsafe_allow_html=True,
        )
        if chungjang:
            st.caption(chungjang)
        if art_no or art_title:
            meta = "  |  ".join(filter(None, [
                f"**{art_no}**" if art_no else "",
                art_title if art_title else "",
            ]))
            st.markdown(meta)
        st.markdown("---")
        if not render_comparison(row, keyword, original_lookup):
            content_hl = highlight(content, keyword)
            st.markdown(
                f'<div style="line-height:1.8;white-space:pre-wrap">{content_hl}</div>',
                unsafe_allow_html=True,
            )


# ── 메인 앱 ──────────────────────────────────────────────────────────────────

def main():
    prevent_browser_translation()

    # ── 인증 ──────────────────────────────────────────────────────────────────
    auth_config = get_auth_config()

    if auth_config is None:
        st.error("⚠️ 인증 설정 파일(config.yaml)이 없습니다.")
        st.stop()

    authenticator = stauth.Authenticate(
        auth_config["credentials"],
        auth_config["cookie"]["name"],
        auth_config["cookie"]["key"],
        auth_config["cookie"]["expiry_days"],
    )

    authenticator.login(location="main", fields={
        "Form name": "건설안전기술사 법 분석 — 로그인",
        "Username": "아이디",
        "Password": "비밀번호",
        "Login": "로그인",
    })

    status = st.session_state.get("authentication_status")

    if status is False:
        st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
        st.stop()
    elif status is None:
        st.stop()

    # ── 로그인 성공 ───────────────────────────────────────────────────────────
    with st.sidebar:
        username = st.session_state.get("username", "")
        st.caption(f"로그인: **{username}**")
        authenticator.logout("로그아웃")

    st.title("⚖️ 건설안전기술사 법령 분석 프로그램")

    data_mtime = os.path.getmtime(DATA_PATH) if os.path.exists(DATA_PATH) else None
    df = load_data(data_mtime)

    if df.empty:
        st.error(
            "데이터 파일(`data/law_data.xlsx`)이 없습니다.  \n"
            "먼저 터미널에서 `python extract_law.py`를 실행해 데이터를 생성하세요."
        )
        st.stop()

    vec, matrix = build_tfidf(df)
    original_lookup = build_original_lookup(df)

    # ── 사이드바 필터 ─────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("🔧 법령 조회 설정")

        search_mode = st.radio(
            "조문 조회 방식",
            ["조문 키워드 검색", "관련 조문 유사도 검색 (TF-IDF)"],
            help="키워드: 검색어가 직접 포함된 조문 / 유사도: 표현이 조금 달라도 관련성이 높은 조문",
        )

        all_codes = sorted(df["법령코드"].unique().tolist())
        code_filter = st.multiselect(
            "대상 법령",
            all_codes,
            default=all_codes,
            format_func=lambda c: LAW_NAMES.get(c, c),
        )

        amendment_filter = st.selectbox(
            "개정 구분",
            ["전체", "개정·신설만", "원본(2024년)만"],
        )

        if search_mode == "관련 조문 유사도 검색 (TF-IDF)":
            threshold = st.slider("최소 유사도 (%)", 5, 80, 15, 5) / 100
            top_n = st.slider("최대 조문 수", 10, 100, 30, 10)
        else:
            threshold = 0.0
            top_n = 300

        st.markdown("---")
        total = len(df)
        st.caption(f"총 조문 수: **{total:,}**개")
        for code in all_codes:
            cnt = len(df[df["법령코드"] == code])
            name = LAW_NAMES.get(code, code)
            st.caption(f"  {name}: {cnt:,}")

    # ── 필터 적용 함수 ────────────────────────────────────────────────────────
    def apply_filters(result_df: pd.DataFrame) -> pd.DataFrame:
        result_df = result_df[result_df["법령코드"].isin(code_filter)]
        if amendment_filter == "개정·신설만":
            result_df = result_df[result_df["개정여부"].isin(["amended", "new"])]
        elif amendment_filter == "원본(2024년)만":
            result_df = result_df[result_df["개정여부"] == "original"]
        return result_df

    # ── 탭 구성 ───────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔍 법령 조문 검색",
        "📋 개정사항 분석",
        "📖 법령별 조문 보기",
        "🛠 데이터 보정",
    ])

    # ── TAB 1: 법령 검색 ──────────────────────────────────────────────────────
    with tab1:
        keyword = st.text_input(
            "🔍 법령·조문 검색어",
            placeholder="예: 비계, 안전난간, 산소결핍, 중대재해, 안전보건관리책임자",
        )

        if not keyword:
            st.info("검색어를 입력하면 건설안전기술사 시험과 관련된 법령 조문을 찾아드립니다.")
        else:
            if search_mode == "조문 키워드 검색":
                result_df = search_keyword(df, keyword)
                result_df = apply_filters(result_df)
                scores = {i: None for i in result_df.index}
            else:
                sim_results = search_similar(df, vec, matrix, keyword,
                                             top_n=top_n, threshold=threshold)
                result_df = apply_filters(df.iloc[[i for i, _ in sim_results]])
                indices = result_df.index.tolist()
                score_map = {i: s for i, s in sim_results}
                scores = {i: score_map.get(i) for i in indices}

            st.markdown(f"### 조문 검색 결과: **{len(result_df)}**개")

            if result_df.empty:
                st.warning("검색된 조문이 없습니다. 다른 법령 용어를 입력하거나 유사도 기준을 낮춰보세요.")
            else:
                # 법령코드별 탭
                codes_in_result = [
                    c for c in all_codes if c in result_df["법령코드"].values
                ]
                if len(codes_in_result) > 1:
                    law_tabs = st.tabs([
                        f"{LAW_NAMES.get(c, c)} ({len(result_df[result_df['법령코드']==c])})"
                        for c in codes_in_result
                    ])
                    for ltab, code in zip(law_tabs, codes_in_result):
                        with ltab:
                            subset = result_df[result_df["법령코드"] == code]
                            for idx, row in subset.iterrows():
                                render_law_card(
                                    row,
                                    keyword=keyword,
                                    score=scores.get(idx),
                                    original_lookup=original_lookup,
                                )
                else:
                    for idx, row in result_df.iterrows():
                        render_law_card(
                            row,
                            keyword=keyword,
                            score=scores.get(idx),
                            original_lookup=original_lookup,
                        )

    # ── TAB 2: 개정 현황 ──────────────────────────────────────────────────────
    with tab2:
        st.subheader("법령 개정사항 분석")

        amended_df = df[df["개정여부"].isin(["amended", "new"])].copy()

        if amended_df.empty:
            st.info("개정·신설 항목이 없습니다. extract_law.py 실행 후 다시 확인하세요.")
        else:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("전체 조문", f"{len(df):,}개")
            with col2:
                st.metric("개정 조문", f"{len(amended_df[amended_df['개정여부']=='amended']):,}개")
            with col3:
                st.metric("신설 조문", f"{len(amended_df[amended_df['개정여부']=='new']):,}개")

            st.markdown("---")

            # 법령별 개정 항목 수
            st.markdown("### 법령별 개정·신설 항목 수")
            amend_count = (
                amended_df.groupby("법령명").size()
                .reset_index(name="항목 수")
                .sort_values("항목 수", ascending=True)
            )
            fig = px.bar(
                amend_count, x="항목 수", y="법령명", orientation="h",
                text="항목 수", color="항목 수",
                color_continuous_scale="Reds",
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(height=400, coloraxis_showscale=False,
                              margin=dict(l=10, r=60))
            st.plotly_chart(fig, use_container_width=True)

            # 시행일별 타임라인
            st.markdown("### 시행일별 법령 개정 현황")
            timeline_df = (
                amended_df[amended_df["시행일"].ne("미상")]
                .groupby(["시행일", "법령명"])
                .size()
                .reset_index(name="조문 수")
                .sort_values("시행일")
            )
            if not timeline_df.empty:
                fig2 = px.bar(
                    timeline_df, x="시행일", y="조문 수", color="법령명",
                    barmode="stack",
                    labels={"시행일": "시행일", "조문 수": "조문 수"},
                )
                fig2.update_layout(height=400)
                st.plotly_chart(fig2, use_container_width=True)

            # 개정 조문 목록
            st.markdown("### 개정·신설 조문 목록")
            filter_type = st.selectbox("개정 유형", ["전체", "개정", "신설"])
            filter_law = st.selectbox(
                "조회 법령",
                ["전체"] + sorted(amended_df["법령명"].unique().tolist()),
            )

            list_df = amended_df.copy()
            if filter_type != "전체":
                type_map = {"개정": "amended", "신설": "new"}
                list_df = list_df[list_df["개정여부"] == type_map[filter_type]]
            if filter_law != "전체":
                list_df = list_df[list_df["법령명"] == filter_law]

            show_df = list_df[["법령명", "조문번호", "조문제목", "시행일", "개정여부",
                                "개정내용요약", "조문내용"]].copy()
            show_df["개정여부"] = show_df["개정여부"].mask(
                list_df.apply(is_new_article, axis=1), "new"
            ).map(
                {"amended": "개정", "new": "신설", "original": "원본"})
            show_df["조문내용"] = show_df["조문내용"].str[:100] + "..."
            st.dataframe(show_df, use_container_width=True, hide_index=True,
                         column_config={
                             "조문내용": st.column_config.TextColumn("조문내용(요약)", width="large"),
                         })

            if not list_df.empty:
                st.markdown("### 개정 상세 비교")
                detail_indices = list_df.index.tolist()

                def format_detail_label(idx: int) -> str:
                    item = list_df.loc[idx]
                    status_label = "신설" if is_new_article(item) else "개정"
                    parts = [
                        status_label,
                        str(item.get("법령명", "")).strip(),
                        str(item.get("조문번호", "")).strip(),
                        str(item.get("조문제목", "")).strip(),
                        str(item.get("시행일", "")).strip(),
                    ]
                    return " | ".join([part for part in parts if part and part != "미상"])

                selected_detail_idx = st.selectbox(
                    "비교할 조문",
                    detail_indices,
                    format_func=format_detail_label,
                )
                render_law_card(
                    list_df.loc[selected_detail_idx],
                    keyword="",
                    score=None,
                    original_lookup=original_lookup,
                )

    # ── TAB 3: 법령별 조문 보기 ───────────────────────────────────────────────
    with tab3:
        st.subheader("법령별 조문 전체 보기")

        law_names_list = sorted(df["법령명"].unique().tolist())
        sel_law = st.selectbox("조회할 법령", law_names_list)

        law_df = df[df["법령명"] == sel_law].copy()

        # 편장절 필터
        chungjang_list = sorted(law_df["편장절"].unique().tolist())
        if len(chungjang_list) > 1:
            sel_chungjang = st.selectbox(
                "편/장/절 선택", ["전체"] + [c for c in chungjang_list if c]
            )
            if sel_chungjang != "전체":
                law_df = law_df[law_df["편장절"] == sel_chungjang]

        st.markdown(f"**{sel_law}** — {len(law_df):,}개 조문")

        amendment_only = st.checkbox("개정·신설 조문만 보기", value=False)
        if amendment_only:
            law_df = law_df[law_df["개정여부"].isin(["amended", "new"])]

        for _, row in law_df.iterrows():
            render_law_card(row, keyword="", score=None, original_lookup=original_lookup)

    # ── TAB 4: 데이터 보정 ────────────────────────────────────────────────────
    with tab4:
        st.subheader("법령명·시행일 보정")
        st.caption(
            "`기타`로 분류되었거나 시행일이 `미상`인 항목을 보정합니다. "
            "저장하면 검색, 법령별 보기, 개정사항 분석에 바로 반영됩니다."
        )

        correction_scope = st.radio(
            "보정 대상",
            ["기타/시행일 미상", "개정·신설 전체", "전체"],
            horizontal=True,
        )

        source_options = ["전체"] + sorted([
            source for source in df.get("출처파일", pd.Series(dtype=str)).astype(str).unique()
            if source
        ])
        source_filter = st.selectbox("출처파일 필터", source_options)

        correction_df = df.copy()
        if correction_scope == "기타/시행일 미상":
            mask = (
                correction_df["법령코드"].eq("기타")
                | correction_df["법령명"].isin(["기타", "법규 요약", ""])
                | correction_df["시행일"].isin(["미상", ""])
            )
            correction_df = correction_df[mask]
        elif correction_scope == "개정·신설 전체":
            correction_df = correction_df[correction_df["개정여부"].isin(["amended", "new"])]

        if source_filter != "전체":
            correction_df = correction_df[correction_df["출처파일"].astype(str).eq(source_filter)]

        st.markdown(f"**보정 대상: {len(correction_df):,}개**")

        if correction_df.empty:
            st.info("현재 조건에 해당하는 보정 대상이 없습니다.")
        else:
            c1, c2 = st.columns([1, 5])
            with c1:
                auto_clicked = st.button("자동 추정 적용", type="secondary")
            with c2:
                st.caption("법령명·법령코드는 파일명/조문내용의 주요 법령명을, 시행일은 파일명/내용의 날짜를 기준으로 추정합니다.")

            if auto_clicked:
                updated_df = df.copy()
                changed = 0
                for idx, row in correction_df.iterrows():
                    new_name, new_code = infer_law_info(row)
                    new_date = infer_enforcement_date(row)
                    new_status = normalize_amendment(row)

                    before = updated_df.loc[idx, ["법령명", "법령코드", "시행일", "개정여부"]].astype(str).tolist()
                    updated_df.loc[idx, "법령명"] = new_name or row.get("법령명", "")
                    updated_df.loc[idx, "법령코드"] = new_code or row.get("법령코드", "")
                    updated_df.loc[idx, "시행일"] = new_date or row.get("시행일", "")
                    updated_df.loc[idx, "개정여부"] = new_status
                    updated_df.loc[idx, "검색텍스트"] = make_search_text(updated_df.loc[idx])
                    after = updated_df.loc[idx, ["법령명", "법령코드", "시행일", "개정여부"]].astype(str).tolist()
                    if before != after:
                        changed += 1

                save_law_data(updated_df)
                st.cache_data.clear()
                st.cache_resource.clear()
                st.success(f"자동 추정으로 {changed:,}개 항목을 보정했습니다.")
                st.rerun()

            editor_df = correction_df.copy()
            editor_df["조문내용 미리보기"] = editor_df["조문내용"].astype(str).str.replace("\n", " ", regex=False).str[:180]
            editor_cols = [
                "id", "법령명", "법령코드", "시행일", "개정여부", "변경유형", "개정내용요약",
                "출처파일", "조문번호", "조문제목", "조문내용 미리보기",
            ]
            editor_df = editor_df[[col for col in editor_cols if col in editor_df.columns]]

            law_code_options = sorted(set(LAW_NAMES.keys()) | set(df["법령코드"].astype(str).unique()))
            edited_df = st.data_editor(
                editor_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["id", "출처파일", "조문번호", "조문제목", "조문내용 미리보기"],
                column_config={
                    "법령코드": st.column_config.SelectboxColumn("법령코드", options=law_code_options),
                    "개정여부": st.column_config.SelectboxColumn(
                        "개정여부",
                        options=["original", "amended", "new"],
                        help="original=원본, amended=개정, new=신설",
                    ),
                    "시행일": st.column_config.TextColumn(
                        "시행일",
                        help="YYYY-MM-DD 형식 권장. 확인 전이면 미상 입력",
                    ),
                    "조문내용 미리보기": st.column_config.TextColumn("조문내용 미리보기", width="large"),
                },
            )

            if st.button("보정 내용 저장", type="primary"):
                updated_df = df.copy()
                changed = 0
                id_to_index = {
                    str(row_id): idx for idx, row_id in updated_df["id"].astype(str).items()
                }

                for _, edited_row in edited_df.iterrows():
                    row_id = str(edited_row.get("id", "")).strip()
                    if row_id not in id_to_index:
                        continue
                    idx = id_to_index[row_id]
                    before = updated_df.loc[idx, EDITABLE_COLUMNS].astype(str).tolist()

                    for col in EDITABLE_COLUMNS:
                        if col in edited_row.index:
                            updated_df.loc[idx, col] = str(edited_row[col]).strip()

                    if "신설" in str(updated_df.loc[idx, "변경유형"]) or "신설" in str(updated_df.loc[idx, "개정내용요약"]):
                        updated_df.loc[idx, "개정여부"] = "new"
                    elif str(updated_df.loc[idx, "개정여부"]).strip() == "":
                        updated_df.loc[idx, "개정여부"] = normalize_amendment(updated_df.loc[idx])

                    updated_df.loc[idx, "검색텍스트"] = make_search_text(updated_df.loc[idx])
                    after = updated_df.loc[idx, EDITABLE_COLUMNS].astype(str).tolist()
                    if before != after:
                        changed += 1

                save_law_data(updated_df)
                st.cache_data.clear()
                st.cache_resource.clear()
                st.success(f"{changed:,}개 항목의 보정 내용을 저장했습니다.")
                st.rerun()


if __name__ == "__main__":
    main()
