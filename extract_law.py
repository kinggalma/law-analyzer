"""
법 분석 프로그램 — 데이터 추출 스크립트
- 실행: python extract_law.py
- 법규 Sumary.pdf + 25년 이후 법개정/*.pdf → data/law_data.xlsx 생성
- 최초 1회만 실행. 이후 추가 파일은 update_law.py로 갱신.
"""

import os
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

try:
    import pdfplumber
except ImportError:
    print("pdfplumber 미설치. 실행: pip install -r requirements-extract.txt")
    sys.exit(1)

# ── 경로 ──────────────────────────────────────────────────────────────────────
BASE_DIR       = Path(__file__).parent
AMENDMENT_DIR  = BASE_DIR / "25년 이후 법개정"
SUMARY_PDF     = BASE_DIR / "법규 Sumary.pdf"
DATA_DIR       = BASE_DIR / "data"
OUTPUT_PATH    = DATA_DIR / "law_data.xlsx"
MANUAL_PATH    = DATA_DIR / "manual_input.xlsx"

DATA_DIR.mkdir(exist_ok=True)

# ── 법령 카테고리 매핑 ─────────────────────────────────────────────────────────
LAW_CODE_MAP = {
    "산업안전보건법": "산안법",
    "산업안전보건기준에 관한 규칙": "산안기준",
    "산업안전보건기준": "산안기준",
    "건설기술진흥법": "건진법",
    "시설물의 안전 및 유지관리에 관한 특별법": "시설물법",
    "시설물안전법": "시설물법",
    "중대재해처벌법": "중처법",
    "중대재해 처벌 등에 관한 법률": "중처법",
    "산업안전보건관리비": "산안관리비",
    "콘크리트공사 표준안전 작업지침": "콘크리트",
    "콘크리트": "콘크리트",
}

# 파일명으로 법령명 추정
FILENAME_LAW_MAP = [
    ("산업안전보건기준에 관한 규칙", "산안기준"),
    ("산업안전보건기준", "산안기준"),
    ("산업안전보건법", "산안법"),
    ("건설기술진흥법", "건진법"),
    ("시설물의 안전 및 유지관리에 관한 특별법", "시설물법"),
    ("시설물 안전법", "시설물법"),
    ("중대재해", "중처법"),
    ("산업안전보건관리비", "산안관리비"),
    ("콘크리트", "콘크리트"),
]

# ── 정규식 패턴 ───────────────────────────────────────────────────────────────
ARTICLE_PAT  = re.compile(r"제\s*(\d+)\s*조(?:의\s*(\d+))?\s*(?:\(([^)]{1,40})\))?")
CHAPTER_PAT  = re.compile(r"제\s*(\d+)\s*장\s*(.{0,40})")
SECTION_PAT  = re.compile(r"제\s*(\d+)\s*절\s*(.{0,40})")
PART_PAT     = re.compile(r"제\s*(\d+)\s*편\s*(.{0,40})")
DATE_PATS    = [
    re.compile(r"(\d{2})\.(\d{1,2})\.(\d{2})\s*시행"),
    re.compile(r"(\d{4})\.(\d{1,2})\.(\d{2})\s*시행"),
    re.compile(r"(\d{2})\.(\d{1,2})\.(\d{2})\s*개정"),
    re.compile(r"(\d{4})\.(\d{1,2})\.(\d{2})\s*개정"),
]
CHUNK_SIZE   = 500  # 조문 파싱 실패 시 글자 수 단위 청크
SUMMARY_OCR_DPI = int(os.environ.get("SUMMARY_OCR_DPI", "170"))
SUMMARY_OCR_MAX_PAGES = int(os.environ.get("SUMMARY_OCR_MAX_PAGES", "0"))


# ── 유틸리티 함수 ─────────────────────────────────────────────────────────────

def detect_pdf_type(path: Path, sample: int = 5) -> str:
    """'text' | 'image' | 'mixed' 반환"""
    try:
        with pdfplumber.open(str(path)) as pdf:
            pages = min(sample, len(pdf.pages))
            chars = sum(len(pdf.pages[i].extract_text() or "") for i in range(pages))
            avg = chars / max(pages, 1)
        if avg >= 200:
            return "text"
        elif avg >= 30:
            return "mixed"
        return "image"
    except Exception:
        return "unknown"


def parse_enforcement_date(filename: str) -> str:
    """파일명에서 시행일/개정일 추출 → 'YYYY-MM-DD'"""
    for pat in DATE_PATS:
        m = pat.search(filename)
        if m:
            y, mo, d = m.groups()
            year = f"20{y}" if len(y) == 2 else y
            return f"{year}-{int(mo):02d}-{int(d):02d}"
    return "미상"


def guess_law_info(filename: str) -> tuple[str, str]:
    """파일명으로 법령명과 법령코드 추정 → (law_name, law_code)"""
    for law_name, law_code in FILENAME_LAW_MAP:
        if law_name in filename:
            return law_name, law_code
    return "기타", "기타"


def infer_law_from_text(text: str) -> tuple[str, str]:
    """조문 텍스트 안의 법령명을 기준으로 법령명/코드를 추정한다."""
    for law_name, law_code in LAW_CODE_MAP.items():
        if law_name in text:
            canonical_name = next(
                (name for name, code in LAW_CODE_MAP.items() if code == law_code and len(name) >= len(law_name)),
                law_name,
            )
            return canonical_name, law_code
    return "법규 Summary", "기타"


def normalize_article_no(value: str) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def normalize_match_text(value: str) -> str:
    value = re.sub(r"\s+", "", str(value or ""))
    return re.sub(r"[^0-9A-Za-z가-힣]", "", value)[:900]


def parse_law_structure(text: str, law_name: str, law_code: str,
                        enforcement_date: str, source_file: str,
                        amendment: str = "original") -> list[dict]:
    """텍스트에서 조문 단위로 분리. 실패 시 청크 단위 폴백."""
    records = []
    current_part = ""
    current_chapter = ""
    current_section = ""
    current_article_no = ""
    current_article_title = ""
    current_content_lines: list[str] = []

    def flush(art_no, art_title, lines):
        content = "\n".join(lines).strip()
        if not content or len(content) < 5:
            return
        records.append({
            "법령명": law_name,
            "법령코드": law_code,
            "편장절": " ".join(filter(None, [current_part, current_chapter, current_section])),
            "조문번호": art_no,
            "조문제목": art_title,
            "조문내용": content[:1000],
            "시행일": enforcement_date,
            "개정여부": amendment,
            "개정내용요약": "",
            "출처파일": source_file,
            "검색텍스트": f"{law_name} {art_no} {art_title} {content}"[:1500],
        })

    lines = text.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # 편/장/절 감지
        pm = PART_PAT.match(stripped)
        if pm:
            current_part = f"제{pm.group(1)}편 {pm.group(2).strip()}"
            continue
        cm = CHAPTER_PAT.match(stripped)
        if cm:
            current_chapter = f"제{cm.group(1)}장 {cm.group(2).strip()}"
            continue
        sm = SECTION_PAT.match(stripped)
        if sm:
            current_section = f"제{sm.group(1)}절 {sm.group(2).strip()}"
            continue

        # 조문 감지
        am = ARTICLE_PAT.match(stripped)
        if am:
            # 이전 조문 저장
            if current_article_no:
                flush(current_article_no, current_article_title, current_content_lines)
            # 새 조문 시작
            num = am.group(1)
            sub = am.group(2) or ""
            title = (am.group(3) or "").strip()
            current_article_no = f"제{num}조" + (f"의{sub}" if sub else "")
            current_article_title = title
            current_content_lines = [stripped]
        else:
            current_content_lines.append(stripped)

    # 마지막 조문 저장
    if current_article_no:
        flush(current_article_no, current_article_title, current_content_lines)

    # 조문이 하나도 파싱되지 않으면 청크 단위로 폴백
    if not records:
        records = chunk_fallback(text, law_name, law_code, enforcement_date, source_file, amendment)

    return records


def chunk_fallback(text: str, law_name: str, law_code: str,
                   enforcement_date: str, source_file: str,
                   amendment: str) -> list[dict]:
    """조문 파싱 실패 시 500자 단위 청크로 분리."""
    records = []
    text = text.strip()
    for i in range(0, len(text), CHUNK_SIZE):
        chunk = text[i:i + CHUNK_SIZE].strip()
        if len(chunk) < 20:
            continue
        chunk_no = i // CHUNK_SIZE + 1
        records.append({
            "법령명": law_name,
            "법령코드": law_code,
            "편장절": "",
            "조문번호": f"섹션{chunk_no}",
            "조문제목": "",
            "조문내용": chunk,
            "시행일": enforcement_date,
            "개정여부": amendment,
            "개정내용요약": "",
            "출처파일": source_file,
            "검색텍스트": f"{law_name} {chunk}"[:1500],
        })
    return records


def parse_comparison_table(path: Path, law_name: str, law_code: str,
                            enforcement_date: str) -> list[dict]:
    """신구조문대비표(2열 테이블 구조) PDF 파싱."""
    records = []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 2:
                            continue
                        current_text = str(row[0] or "").strip()
                        amended_text = str(row[1] or "").strip()
                        if not current_text or not amended_text:
                            continue
                        # 헤더 행 스킵
                        if current_text in ("현행", "개정안", "현  행", "개  정  안"):
                            continue
                        if len(amended_text) < 5:
                            continue
                        is_new = any(token in current_text for token in ("신설", "<신 설>", "<신설>"))
                        # 조문번호 추출
                        am = ARTICLE_PAT.search(amended_text)
                        art_no = ""
                        if am:
                            num = am.group(1)
                            sub = am.group(2) or ""
                            art_no = f"제{num}조" + (f"의{sub}" if sub else "")
                        records.append({
                            "법령명": law_name,
                            "법령코드": law_code,
                            "편장절": "",
                            "조문번호": art_no,
                            "조문제목": "",
                            "조문내용": amended_text[:1000],
                            "시행일": enforcement_date,
                            "개정여부": "new" if is_new else "amended",
                            "개정내용요약": "신규 추가 조문" if is_new else "기존 조문 변경",
                            "개정전조문내용": "" if is_new else current_text[:1000],
                            "변경유형": "신설" if is_new else "개정",
                            "출처파일": path.name,
                            "검색텍스트": f"{law_name} {art_no} {amended_text}"[:1500],
                        })
    except Exception as e:
        print(f"    [테이블 파싱 오류] {path.name}: {e}")
    return records


def extract_pdf(path: Path, law_name: str, law_code: str,
                enforcement_date: str, amendment: str = "original") -> list[dict]:
    """PDF 파일에서 법령 데이터를 추출한다."""
    pdf_type = detect_pdf_type(path)
    print(f"    유형: {pdf_type}  |  {path.name}")

    if pdf_type == "image":
        print(f"    → 이미지 스캔. 건너뜁니다. (수동 입력: data/manual_input.xlsx)")
        return []

    records = []
    # 신구조문대비표는 테이블 파싱 우선
    if "신구조문대비표" in path.name or "신구" in path.name:
        table_records = parse_comparison_table(path, law_name, law_code, enforcement_date)
        if table_records:
            print(f"    → 신구조문대비표 파싱: {len(table_records)}개 항목")
            return table_records

    try:
        with pdfplumber.open(str(path)) as pdf:
            all_text_parts = []
            for page in pdf.pages:
                t = page.extract_text() or ""
                if t.strip():
                    all_text_parts.append(t)
            full_text = "\n".join(all_text_parts)
    except Exception as e:
        print(f"    → 텍스트 추출 실패: {e}")
        return []

    if not full_text.strip():
        print(f"    → 텍스트 없음. 건너뜁니다.")
        return []

    records = parse_law_structure(
        full_text, law_name, law_code, enforcement_date, path.name, amendment
    )
    print(f"    → 추출: {len(records)}개 항목")
    return records


def extract_image_pdf_text_with_ocr(path: Path) -> str:
    """이미지 스캔 PDF를 EasyOCR로 텍스트화한다."""
    try:
        import fitz
        import easyocr
        import numpy as np
    except ImportError as exc:
        print(f"    → OCR 의존성 없음: {exc}")
        print("       설치: pip install easyocr pymupdf")
        return ""

    try:
        reader = easyocr.Reader(["ko", "en"], gpu=False)
        doc = fitz.open(str(path))
        max_pages = SUMMARY_OCR_MAX_PAGES or len(doc)
        max_pages = min(max_pages, len(doc))
        scale = SUMMARY_OCR_DPI / 72
        matrix = fitz.Matrix(scale, scale)
        text_parts = []
        print(f"    → OCR 시작: {max_pages}/{len(doc)}쪽, {SUMMARY_OCR_DPI}dpi")
        for page_idx in range(max_pages):
            page = doc.load_page(page_idx)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            lines = reader.readtext(image, detail=0, paragraph=True)
            page_text = "\n".join(str(line) for line in lines if str(line).strip())
            if page_text.strip():
                text_parts.append(page_text)
            print(f"      OCR {page_idx + 1}/{max_pages}쪽: {len(page_text):,}자")
        doc.close()
        return "\n".join(text_parts)
    except Exception as exc:
        print(f"    → OCR 실패: {exc}")
        return ""


def extract_summary_pdf(path: Path) -> list[dict]:
    """법규 Sumary.pdf를 기준 조문 데이터로 추출한다."""
    pdf_type = detect_pdf_type(path, sample=10)
    print(f"  유형: {pdf_type}")

    if pdf_type in ("text", "mixed"):
        records = extract_pdf(path, "법규 Summary", "기타", "2024-12-31", "original")
    elif pdf_type == "image":
        full_text = extract_image_pdf_text_with_ocr(path)
        if not full_text.strip():
            return []
        records = parse_law_structure(
            full_text, "법규 Summary", "기타", "2024-12-31", path.name, "original"
        )
        print(f"    → OCR 추출: {len(records)}개 항목")
    else:
        print("  → 알 수 없는 PDF 유형입니다. 건너뜁니다.")
        return []

    for record in records:
        text = " ".join([
            str(record.get("법령명", "")),
            str(record.get("조문제목", "")),
            str(record.get("조문내용", "")),
        ])
        law_name, law_code = infer_law_from_text(text)
        if law_code != "기타":
            record["법령명"] = law_name
            record["법령코드"] = law_code
        record["출처파일"] = path.name
        record["개정여부"] = "original"
        record["검색텍스트"] = (
            f"{record.get('법령명', '')} {record.get('조문번호', '')} "
            f"{record.get('조문제목', '')} {record.get('조문내용', '')}"
        )[:1500]
    return records


def link_amendments_to_summary(df: pd.DataFrame) -> pd.DataFrame:
    """개정 행을 Summary 기준 조문에 연결한다."""
    df = df.copy()
    if "기준조문ID" not in df.columns:
        df["기준조문ID"] = ""
    if "기준연결방식" not in df.columns:
        df["기준연결방식"] = ""

    summary_mask = df["개정여부"].eq("original") & df["출처파일"].astype(str).str.contains("법규 Sumary.pdf|법규 Summary.pdf", case=False, regex=True, na=False)
    summary_df = df[summary_mask].copy()
    if summary_df.empty:
        print("  → Summary 기준 조문이 없어 개정사항 연결을 건너뜁니다.")
        return df

    exact_lookup: dict[tuple[str, str], int] = {}
    summary_by_law: dict[str, list[tuple[int, str]]] = {}
    for idx, row in summary_df.iterrows():
        base_id = str(row.get("id", "")).strip()
        law = str(row.get("법령명", "")).strip()
        article = normalize_article_no(row.get("조문번호", ""))
        if law and article and (law, article) not in exact_lookup:
            exact_lookup[(law, article)] = idx
        match_text = normalize_match_text(f"{row.get('조문제목', '')} {row.get('조문내용', '')}")
        summary_by_law.setdefault(law, []).append((idx, match_text))
        df.loc[idx, "기준조문ID"] = base_id
        df.loc[idx, "기준연결방식"] = "summary"

    linked = 0
    amendment_mask = df["개정여부"].isin(["amended", "new"])
    for idx, row in df[amendment_mask].iterrows():
        law = str(row.get("법령명", "")).strip()
        article = normalize_article_no(row.get("조문번호", ""))
        exact_idx = exact_lookup.get((law, article))
        if exact_idx is not None:
            df.loc[idx, "기준조문ID"] = str(df.loc[exact_idx, "id"])
            df.loc[idx, "기준연결방식"] = "법령명+조문번호"
            linked += 1
            continue

        candidates = summary_by_law.get(law, [])
        target_text = normalize_match_text(f"{row.get('조문제목', '')} {row.get('조문내용', '')}")
        best_idx = None
        best_score = 0.0
        if target_text:
            for cand_idx, cand_text in candidates:
                if not cand_text:
                    continue
                score = SequenceMatcher(None, target_text, cand_text).ratio()
                if score > best_score:
                    best_idx = cand_idx
                    best_score = score

        if best_idx is not None and best_score >= 0.35:
            df.loc[idx, "기준조문ID"] = str(df.loc[best_idx, "id"])
            df.loc[idx, "기준연결방식"] = f"내용유사도 {best_score:.2f}"
            linked += 1
        else:
            df.loc[idx, "기준연결방식"] = "미연결(신규 가능)"

    print(f"  → Summary 기준 연결: {linked:,}/{int(amendment_mask.sum()):,}개 개정·신설 항목")
    return df


def try_hwp_extraction(path: Path) -> str:
    """hwp5txt로 HWP 텍스트 추출 시도."""
    try:
        import olefile
        ole = olefile.OleFileIO(str(path))
        if ole.exists("PrvText"):
            raw = ole.openstream("PrvText").read()
            text = raw.decode("utf-16-le", errors="ignore")
            return text
    except Exception:
        pass
    return ""


def load_manual_input() -> list[dict]:
    """data/manual_input.xlsx 가 있으면 로드한다."""
    if not MANUAL_PATH.exists():
        return []
    try:
        df = pd.read_excel(MANUAL_PATH)
        return df.to_dict("records")
    except Exception as e:
        print(f"manual_input.xlsx 로드 실패: {e}")
        return []


def assign_ids(records: list[dict]) -> list[dict]:
    """중복 없는 id 부여."""
    for i, r in enumerate(records, start=1):
        r["id"] = i
    return records


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    all_records: list[dict] = []
    history_records: list[dict] = []

    # ── 1. 법규 Sumary.pdf를 기준 데이터로 먼저 처리 ─────────────────────────
    if SUMARY_PDF.exists():
        print(f"\n[법규 Sumary.pdf] 기준 데이터 처리 중...")
        records = extract_summary_pdf(SUMARY_PDF)
        print(f"  → 법규 Sumary 기준 추출: {len(records)}개 항목")
        all_records.extend(records)
    else:
        print(f"\n⚠️  법규 Sumary.pdf 없음: {SUMARY_PDF}")

    # ── 2. 25년 이후 법개정 PDF 처리 ─────────────────────────────────────────
    if AMENDMENT_DIR.exists():
        pdf_files = sorted(AMENDMENT_DIR.glob("*.pdf"))
        print(f"\n[25년 이후 법개정] PDF {len(pdf_files)}개 처리 시작...")
        for pdf_path in pdf_files:
            enforcement_date = parse_enforcement_date(pdf_path.name)
            law_name, law_code = guess_law_info(pdf_path.name)
            print(f"\n  처리 중: {pdf_path.name}")
            print(f"    법령명: {law_name}  |  시행일: {enforcement_date}")
            records = extract_pdf(pdf_path, law_name, law_code, enforcement_date, amendment="amended")
            all_records.extend(records)

        # HWP 파일 처리
        hwp_files = sorted(AMENDMENT_DIR.glob("*.hwp"))
        if hwp_files:
            print(f"\n  HWP 파일 {len(hwp_files)}개 처리 중...")
            for hwp_path in hwp_files:
                print(f"\n  처리 중: {hwp_path.name}")
                text = try_hwp_extraction(hwp_path)
                if text.strip():
                    law_name, law_code = guess_law_info(hwp_path.name)
                    enforcement_date = parse_enforcement_date(hwp_path.name)
                    records = parse_law_structure(
                        text, law_name, law_code, enforcement_date, hwp_path.name, "amended"
                    )
                    print(f"    → HWP 추출: {len(records)}개 항목")
                    all_records.extend(records)
                else:
                    print(f"    → HWP 추출 실패. data/manual_input.xlsx에 직접 입력하세요.")
    else:
        print(f"\n⚠️  '25년 이후 법개정' 폴더 없음: {AMENDMENT_DIR}")

    # ── 3. 수동 입력 데이터 병합 ──────────────────────────────────────────────
    manual = load_manual_input()
    if manual:
        print(f"\n[manual_input.xlsx] {len(manual)}개 항목 로드")
        all_records.extend(manual)

    if not all_records:
        print("\n❌ 추출된 데이터가 없습니다. PDF 파일과 경로를 확인하세요.")
        return

    # ── 4. 데이터프레임 생성 및 정제 ─────────────────────────────────────────
    df = pd.DataFrame(all_records)

    # 필수 컬럼 보장
    required_cols = [
        "id", "법령명", "법령코드", "편장절", "조문번호", "조문제목",
        "조문내용", "시행일", "개정여부", "개정내용요약",
        "개정전조문내용", "변경유형", "출처파일", "기준조문ID", "기준연결방식", "검색텍스트"
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    df = df.fillna("")
    df = df[df["조문내용"].str.strip().ne("")]  # 빈 내용 제거

    # id 재부여
    df = df.reset_index(drop=True)
    df["id"] = df.index + 1
    df["id"] = df["id"].astype(str)

    df = link_amendments_to_summary(df)

    # 컬럼 순서 정렬
    df = df[required_cols]

    # ── 5. Excel 저장 ─────────────────────────────────────────────────────────
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="법령조문", index=False)
        # 개정이력 시트 (현재는 빈 템플릿)
        history_cols = ["법령명", "조문번호", "시행일", "개정분류", "개정전", "개정후", "비고"]
        history_df = pd.DataFrame(columns=history_cols)
        history_df.to_excel(writer, sheet_name="개정이력", index=False)

    print(f"\n[완료]")
    print(f"   총 {len(df):,}개 항목 -> {OUTPUT_PATH}")
    print(f"\n   법령별 분포:")
    for code, cnt in df["법령코드"].value_counts().items():
        print(f"     {code}: {cnt:,}개")
    print(f"\n   다음 단계: streamlit run app.py")


if __name__ == "__main__":
    main()
