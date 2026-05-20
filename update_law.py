"""
법 분석 프로그램 — 법령 데이터 갱신 스크립트
- 실행: python update_law.py
- '25년 이후 법개정/' 폴더에 새 PDF를 추가한 뒤 실행하면
  data/law_data.xlsx에 미처리 파일만 추가 추출해 반영한다.
"""

import os
import sys
from pathlib import Path

import pandas as pd

# extract_law 모듈에서 공통 함수 재사용
try:
    from extract_law import (
        AMENDMENT_DIR,
        DATA_DIR,
        OUTPUT_PATH,
        detect_pdf_type,
        extract_pdf,
        guess_law_info,
        parse_enforcement_date,
        try_hwp_extraction,
        parse_law_structure,
    )
except ImportError as e:
    print(f"extract_law.py를 로드할 수 없습니다: {e}")
    print("같은 폴더에 extract_law.py가 있는지 확인하세요.")
    sys.exit(1)

SHEET_NAME = "법령조문"
HISTORY_SHEET = "개정이력"


def get_already_processed(df: pd.DataFrame) -> set[str]:
    """현재 Excel에 이미 들어간 출처파일 목록을 반환한다."""
    if "출처파일" not in df.columns:
        return set()
    return set(df["출처파일"].dropna().astype(str).unique())


def load_existing() -> tuple[pd.DataFrame, pd.DataFrame]:
    """기존 law_data.xlsx 로드. 없으면 빈 DataFrame 반환."""
    if not OUTPUT_PATH.exists():
        print(f"[주의] {OUTPUT_PATH} 파일이 없습니다.")
        print("   먼저 python extract_law.py 를 실행하세요.")
        sys.exit(1)

    df_main = pd.read_excel(OUTPUT_PATH, sheet_name=SHEET_NAME)
    try:
        df_history = pd.read_excel(OUTPUT_PATH, sheet_name=HISTORY_SHEET)
    except Exception:
        df_history = pd.DataFrame()
    return df_main, df_history


def save_excel(df_main: pd.DataFrame, df_history: pd.DataFrame):
    """law_data.xlsx 저장."""
    DATA_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        df_main.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        if df_history is not None and not df_history.empty:
            df_history.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)
        else:
            # 빈 개정이력 시트 유지
            history_cols = ["법령명", "조문번호", "시행일", "개정분류", "개정전", "개정후", "비고"]
            pd.DataFrame(columns=history_cols).to_excel(
                writer, sheet_name=HISTORY_SHEET, index=False)


def reassign_ids(df: pd.DataFrame) -> pd.DataFrame:
    df = df.reset_index(drop=True)
    df["id"] = df.index + 1
    return df


def main():
    print("=" * 60)
    print("  법령 데이터 갱신 스크립트 (update_law.py)")
    print("=" * 60)

    df_main, df_history = load_existing()
    already_processed = get_already_processed(df_main)
    print(f"\n현재 데이터: {len(df_main):,}개 항목, 처리된 파일: {len(already_processed)}개")

    new_records: list[dict] = []

    # ── PDF 신규 파일 처리 ────────────────────────────────────────────────────
    if AMENDMENT_DIR.exists():
        pdf_files = sorted(AMENDMENT_DIR.glob("*.pdf"))
        new_pdfs = [p for p in pdf_files if p.name not in already_processed]
        skip_pdfs = [p for p in pdf_files if p.name in already_processed]

        print(f"\n[25년 이후 법개정] 전체 {len(pdf_files)}개 PDF")
        print(f"  이미 처리됨: {len(skip_pdfs)}개  |  신규: {len(new_pdfs)}개")

        for pdf_path in new_pdfs:
            enforcement_date = parse_enforcement_date(pdf_path.name)
            law_name, law_code = guess_law_info(pdf_path.name)
            print(f"\n  처리 중: {pdf_path.name}")
            print(f"    법령명: {law_name}  |  시행일: {enforcement_date}")
            records = extract_pdf(pdf_path, law_name, law_code, enforcement_date, amendment="amended")
            new_records.extend(records)

        # HWP 신규 파일 처리
        hwp_files = sorted(AMENDMENT_DIR.glob("*.hwp"))
        new_hwps = [p for p in hwp_files if p.name not in already_processed]
        for hwp_path in new_hwps:
            print(f"\n  처리 중 (HWP): {hwp_path.name}")
            text = try_hwp_extraction(hwp_path)
            if text.strip():
                law_name, law_code = guess_law_info(hwp_path.name)
                enforcement_date = parse_enforcement_date(hwp_path.name)
                records = parse_law_structure(
                    text, law_name, law_code, enforcement_date, hwp_path.name, "amended"
                )
                print(f"    → HWP 추출: {len(records)}개 항목")
                new_records.extend(records)
            else:
                print(f"    → HWP 추출 실패. data/manual_input.xlsx에 직접 입력하세요.")
    else:
        print(f"\n[주의] '25년 이후 법개정' 폴더 없음: {AMENDMENT_DIR}")

    if not new_records:
        print("\n[완료] 새로운 파일이 없습니다. law_data.xlsx는 최신 상태입니다.")
        return

    # ── 병합 및 저장 ──────────────────────────────────────────────────────────
    df_new = pd.DataFrame(new_records)

    # 필수 컬럼 보장
    required_cols = [
        "id", "법령명", "법령코드", "편장절", "조문번호", "조문제목",
        "조문내용", "시행일", "개정여부", "개정내용요약",
        "개정전조문내용", "변경유형", "출처파일", "검색텍스트"
    ]
    for col in required_cols:
        if col not in df_new.columns:
            df_new[col] = ""
    for col in required_cols:
        if col not in df_main.columns:
            df_main[col] = ""

    df_merged = pd.concat([df_main, df_new[required_cols]], ignore_index=True)
    df_merged = df_merged.fillna("")
    df_merged = df_merged[df_merged["조문내용"].str.strip().ne("")]
    df_merged = reassign_ids(df_merged)

    save_excel(df_merged, df_history)

    print(f"\n[완료] 갱신 완료!")
    print(f"   추가: {len(df_new):,}개 항목")
    print(f"   총계: {len(df_merged):,}개 항목 -> {OUTPUT_PATH}")
    print(f"\n   법령별 분포:")
    for code, cnt in df_merged["법령코드"].value_counts().items():
        print(f"     {code}: {cnt:,}개")


if __name__ == "__main__":
    main()
