"""
tests/test_ai_builders_output.py
────────────────────────────────────────────────────────
5-4-D: build_slide_2/5/7 실제 PPTX 생성 테스트
천안 부성2지구 PDF → Claude (캐시) → 3개 빌더 → PPTX 저장.
────────────────────────────────────────────────────────
"""

import sys
import os
import traceback

# UTF-8 출력 (Windows cp949 깨짐 방지)
sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)
sys.stderr = open(sys.stderr.fileno(), mode="w", encoding="utf-8", buffering=1)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
OUTPUT   = os.path.join(ROOT, "test_ai_builders_output.pptx")
BIZ_NAME = "천안 부성2지구 도시개발사업"


def extract_pdf_text(pdf_path: str) -> str:
    import fitz
    doc = fitz.open(pdf_path)
    text = "\n".join(page.get_text() for page in doc)
    doc.close()
    return text


def slide_preview(slide) -> list[str]:
    """슬라이드 내 텍스트 프레임에서 앞 50자씩 수집."""
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t:
                out.append(f"  [{sh.shape_type}] {t[:50]}")
    return out


def main():
    print("=" * 70)
    print("5-4-D: AI 빌더 PPTX 생성 테스트")
    print("=" * 70)

    # ── API 키 확인 ──────────────────────────────────────────
    try:
        from modules.claude_api import get_client
        get_client()
        print("[OK] API 클라이언트 초기화")
    except Exception as exc:
        print(f"[FAIL] API 키 없음: {exc}")
        sys.exit(1)

    # ── PDF 추출 ─────────────────────────────────────────────
    if not os.path.exists(PDF_PATH):
        print(f"[FAIL] PDF 없음: {PDF_PATH}")
        sys.exit(1)

    print(f"\n[STEP 1] PDF 텍스트 추출")
    try:
        pdf_text = extract_pdf_text(PDF_PATH)
        print(f"  완료 — {len(pdf_text):,} 자")
    except Exception:
        print(f"[FAIL]\n{traceback.format_exc()}")
        sys.exit(1)

    # ── Claude API (캐시 활용) ────────────────────────────────
    from modules.ai_slide_builders import (
        generate_executive_summary,
        generate_sasae_overview,
        generate_investment_structure,
        build_slide_2_executive_summary,
        build_slide_5_sasae_overview,
        build_slide_7_investment_structure,
    )

    print(f"\n[STEP 2] Claude API 호출 (캐시 활용 시 추가 비용 없음)")
    total_input = 0
    total_output = 0

    try:
        r2 = generate_executive_summary(pdf_text)
        r5 = generate_sasae_overview(pdf_text)
        r7 = generate_investment_structure(pdf_text)
    except Exception:
        print(f"[FAIL]\n{traceback.format_exc()}")
        sys.exit(1)

    for label, r in [("슬라이드2", r2), ("슬라이드5", r5), ("슬라이드7", r7)]:
        if not r.get("ok"):
            print(f"[FAIL] {label}: {r.get('error')}")
            sys.exit(1)
        cached_str = "캐시HIT ✓" if r.get("cached") else "신규API"
        u = r["usage"]
        total_input  += u.get("input_tokens",  0)
        total_output += u.get("output_tokens", 0)
        print(f"  {label}: {cached_str} | in={u['input_tokens']:,} out={u['output_tokens']:,}")

    print(f"\n  토큰 합산: in={total_input:,} / out={total_output:,}")
    print(f"  캐시 호출은 과금 대상 아님 (캐시 저장 시 과금)")

    # ── PPT 빌더 ─────────────────────────────────────────────
    from modules.page_builders import create_presentation_from_template, finalize_presentation
    from modules.ppt_generator import save_presentation

    print(f"\n[STEP 3] PPT 빌더 실행")
    try:
        prs = create_presentation_from_template()
        n   = len(prs.slides)
        print(f"  템플릿 슬라이드 수: {n}")

        # ── 슬라이드 2 ──
        print("=" * 60)
        print(f"[FORCE-DEBUG-BUILD] build_slide_2 호출 - data keys: {list(r2['data'].keys())}")
        print("=" * 60)
        s2 = build_slide_2_executive_summary(
            prs, r2["data"], business_name=BIZ_NAME, page_num=2
        )

        # ── 슬라이드 5 ──
        print("=" * 60)
        print(f"[FORCE-DEBUG-BUILD] build_slide_5 호출 - data keys: {list(r5['data'].keys())}")
        print("=" * 60)
        s5 = build_slide_5_sasae_overview(
            prs, r5["data"], business_name=BIZ_NAME, page_num=6
        )

        # ── 슬라이드 7 ──
        print("=" * 60)
        print(f"[FORCE-DEBUG-BUILD] build_slide_7 호출 - data keys: {list(r7['data'].keys())}")
        print("=" * 60)
        s7 = build_slide_7_investment_structure(
            prs, r7["data"], business_name=BIZ_NAME, page_num=8
        )

        finalize_presentation(prs, n)
        save_presentation(prs, OUTPUT)

    except Exception:
        print(f"[FAIL]\n{traceback.format_exc()}")
        sys.exit(1)

    # ── 검증 출력 ─────────────────────────────────────────────
    print(f"\n[STEP 4] 생성 결과 검증")
    print(f"  파일 경로: {OUTPUT}")
    print(f"  슬라이드 수: {len(prs.slides)}")

    for label, slide in [("슬라이드2 (Executive Summary)", s2),
                          ("슬라이드5 (사모사채 개요)",      s5),
                          ("슬라이드7 (투자구조도)",          s7)]:
        shapes = slide.shapes
        texts  = slide_preview(slide)
        print(f"\n  ── {label} ──")
        print(f"     shape 수: {len(shapes)}")
        print(f"     텍스트 미리보기:")
        if texts:
            for t in texts:
                print(f"     {t}")
        else:
            print("     (텍스트 없음)")

    # ── 문제 감지 ─────────────────────────────────────────────
    issues = []
    for label, slide in [("슬라이드2", s2), ("슬라이드5", s5), ("슬라이드7", s7)]:
        empty_tf = [sh for sh in slide.shapes
                    if sh.has_text_frame and not sh.text_frame.text.strip()]
        if empty_tf:
            issues.append(f"{label}: 빈 텍스트 프레임 {len(empty_tf)}개")

    print(f"\n  ── 발견된 문제 ──")
    if issues:
        for iss in issues:
            print(f"  [경고] {iss}")
    else:
        print("  없음")

    print(f"\n{'='*70}")
    print(f"5-4-D 완료 — {OUTPUT} 파일을 열어 육안 확인하세요.")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
