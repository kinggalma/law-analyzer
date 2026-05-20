"""
PDF 텍스트 추출 가능 여부 확인 스크립트
- 실행: python check_pdf.py
- 법규 Sumary.pdf 및 법개정 PDF 샘플을 검사해 텍스트 기반인지 이미지 스캔인지 확인한다.
"""

import os
import sys

try:
    import pdfplumber
except ImportError:
    print("pdfplumber가 설치되지 않았습니다. 먼저 실행하세요:")
    print("  pip install pdfplumber")
    sys.exit(1)


def check_pdf(path: str, sample_pages: int = 5) -> dict:
    """PDF 파일의 텍스트 추출 품질을 확인한다."""
    result = {
        "path": path,
        "filename": os.path.basename(path),
        "total_pages": 0,
        "char_per_page": 0,
        "type": "unknown",
        "sample_text": "",
        "error": None,
    }
    try:
        with pdfplumber.open(path) as pdf:
            result["total_pages"] = len(pdf.pages)
            pages_to_check = min(sample_pages, len(pdf.pages))
            texts = []
            for i in range(pages_to_check):
                t = pdf.pages[i].extract_text() or ""
                texts.append(t)
            total_chars = sum(len(t) for t in texts)
            result["char_per_page"] = total_chars / max(pages_to_check, 1)
            result["sample_text"] = (texts[0][:200] if texts else "").strip()

            if result["char_per_page"] >= 200:
                result["type"] = "text"
            elif result["char_per_page"] >= 30:
                result["type"] = "mixed"
            else:
                result["type"] = "image"
    except Exception as e:
        result["error"] = str(e)
    return result


def print_result(r: dict):
    status_icon = {"text": "✅", "mixed": "⚠️", "image": "❌", "unknown": "❓"}.get(r["type"], "❓")
    print(f"\n{status_icon} [{r['type'].upper()}] {r['filename']}")
    print(f"   총 페이지: {r['total_pages']}  |  페이지당 평균 {r['char_per_page']:.0f}자")
    if r["error"]:
        print(f"   오류: {r['error']}")
    elif r["sample_text"]:
        preview = r["sample_text"].replace("\n", " ")[:120]
        print(f"   샘플: {preview!r}")


BASE_DIR = os.path.dirname(__file__)
AMENDMENT_DIR = os.path.join(BASE_DIR, "25년 이후 법개정")


def main():
    print("=" * 60)
    print("  PDF 텍스트 추출 가능 여부 확인")
    print("=" * 60)

    # 법규 Sumary.pdf 확인
    sumary_path = os.path.join(BASE_DIR, "법규 Sumary.pdf")
    if os.path.exists(sumary_path):
        print("\n[법규 Sumary.pdf]")
        r = check_pdf(sumary_path, sample_pages=5)
        print_result(r)
        if r["type"] == "text":
            print("   → pdfplumber로 전체 추출 가능합니다 (extract_law.py에서 처리)")
        elif r["type"] == "image":
            print("   → 이미지 스캔입니다. extract_law.py 실행 시 Sumary.pdf는 건너뜁니다.")
            print("     (OCR이 필요하면 별도 협의)")
        else:
            print("   → 혼합형입니다. 텍스트 페이지만 추출하고 이미지 페이지는 건너뜁니다.")
    else:
        print(f"\n⚠️  법규 Sumary.pdf 파일을 찾을 수 없습니다: {sumary_path}")

    # 법개정 PDF 샘플 10개 확인
    if os.path.exists(AMENDMENT_DIR):
        pdfs = sorted([
            os.path.join(AMENDMENT_DIR, f)
            for f in os.listdir(AMENDMENT_DIR)
            if f.lower().endswith(".pdf")
        ])
        print(f"\n[25년 이후 법개정 폴더 — PDF {len(pdfs)}개 중 10개 샘플]")
        for path in pdfs[:10]:
            r = check_pdf(path, sample_pages=3)
            print_result(r)
    else:
        print(f"\n⚠️  '25년 이후 법개정' 폴더를 찾을 수 없습니다: {AMENDMENT_DIR}")

    print("\n" + "=" * 60)
    print("  결론 확인 후 extract_law.py를 실행하세요:")
    print("    python extract_law.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
