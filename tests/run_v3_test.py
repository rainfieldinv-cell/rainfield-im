"""
run_v3_test.py — test_ai_builders_v3.pptx 생성 + 검증
"""
import sys
import os
import io
import traceback

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
OUTPUT   = os.path.join(ROOT, "test_ai_builders_v3.pptx")
BIZ_NAME = "천안 부성2지구 도시개발사업"


def main():
    print("=" * 70)
    print("STEP 3: test_ai_builders_v3.pptx 생성")
    print("=" * 70)

    # PDF 추출
    import fitz
    doc = fitz.open(PDF_PATH)
    pdf_text = "\n".join(page.get_text() for page in doc)
    doc.close()
    print(f"PDF 추출: {len(pdf_text):,} 자")

    # Claude API (캐시 활용)
    from modules.ai_slide_builders import (
        generate_executive_summary,
        generate_sasae_overview,
        generate_investment_structure,
        build_slide_2_executive_summary,
        build_slide_5_sasae_overview,
        build_slide_7_investment_structure,
    )
    from modules.page_builders import create_presentation_from_template, finalize_presentation
    from modules.ppt_generator import save_presentation

    r2 = generate_executive_summary(pdf_text)
    r5 = generate_sasae_overview(pdf_text)
    r7 = generate_investment_structure(pdf_text)

    for label, r in [("S2", r2), ("S5", r5), ("S7", r7)]:
        cached = "캐시HIT" if r.get("cached") else "API"
        ok = "OK" if r.get("ok") else "FAIL"
        err = r.get("error", "")
        print(f"  {label}: {ok} {cached} {err}")

    if not all(r.get("ok") for r in [r2, r5, r7]):
        print("[FAIL] API 호출 실패")
        sys.exit(1)

    # PPT 빌더
    prs = create_presentation_from_template()
    n = len(prs.slides)
    print(f"템플릿 슬라이드: {n}")

    s2 = build_slide_2_executive_summary(prs, r2["data"], business_name=BIZ_NAME, page_num=2)
    s5 = build_slide_5_sasae_overview(prs, r5["data"], business_name=BIZ_NAME, page_num=6)
    s7 = build_slide_7_investment_structure(prs, r7["data"], business_name=BIZ_NAME, page_num=8)

    finalize_presentation(prs, n)
    save_presentation(prs, OUTPUT)
    print(f"\n저장 완료: {OUTPUT}")
    print(f"총 슬라이드: {len(prs.slides)}")

    # ── 슬라이드 2 검증 ──
    print("\n=== 슬라이드 2 (Executive Summary) 검증 ===")
    for i, sh in enumerate(s2.shapes):
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t:
                # 폰트 정보 확인
                fonts = []
                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        if run.text.strip():
                            bold = run.font.bold
                            sz = run.font.size
                            fonts.append(f"bold={bold} sz={sz}")
                            break
                    if fonts:
                        break
                font_info = fonts[0] if fonts else "no-runs"
                print(f"  [{i:02d}] {t[:50]}  | {font_info}")

    # ── 슬라이드 7 상세 검증 ──
    print("\n=== 슬라이드 7 (투자구조도) 상세 검증 ===")
    for i, sh in enumerate(s7.shapes):
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t:
                print(f"  [{i:02d}] TEXT: {t[:60]}")
            else:
                print(f"  [{i:02d}] TEXT: (비어있음)")
        elif sh.shape_type == 19:
            tbl = sh.table
            cells = []
            for r in range(len(tbl.rows)):
                for c in range(len(tbl.columns)):
                    ct = tbl.cell(r, c).text.strip()[:25]
                    cells.append(ct if ct else "(빈)")
            print(f"  [{i:02d}] TABLE {len(tbl.rows)}r x {len(tbl.columns)}c: {' / '.join(cells)}")

    # ── PDF 원본 일치 확인 ──
    print("\n=== PDF 원본 구조도 일치 체크 ===")
    checks = []

    # 1. 신탁사 박스 기관명
    for sh in s7.shapes:
        if sh.shape_type == 19:
            tbl = sh.table
            if len(tbl.rows) == 2 and len(tbl.columns) == 1:
                r0 = tbl.cell(0, 0).text.strip()
                r1 = tbl.cell(1, 0).text.strip()
                if "신탁" in r0:
                    checks.append(("신탁사 기관명", r1, "신한자산신탁" in r1))
                elif "차주" in r0 or "시행" in r0:
                    checks.append(("차주 기관명", r1, "더함도시개발" in r1))
                elif "시공" in r0:
                    checks.append(("시공사 기관명", r1, r1 != ""))
                elif r0 == "" and r1 == "":
                    checks.append(("자산관리자 박스 비움", f"r0=[{r0}] r1=[{r1}]", True))

    # 2. 대주 박스 (4r) - 금액 없어야 함
    for sh in s7.shapes:
        if sh.shape_type == 19:
            tbl = sh.table
            if len(tbl.rows) == 4:
                full = " ".join(tbl.cell(r, 0).text for r in range(len(tbl.rows)))
                no_amount = "억" not in full
                checks.append(("대주 박스 금액 없음", full[:50], no_amount))

    # 3. 투자자 박스 (7r) - 전부 비어야 함
    for sh in s7.shapes:
        if sh.shape_type == 19:
            tbl = sh.table
            if len(tbl.rows) == 7:
                all_empty = all(tbl.cell(r, 0).text.strip() == "" for r in range(len(tbl.rows)))
                checks.append(("투자자 박스 전체 비움", f"all_empty={all_empty}", all_empty))

    # 4. 화살표 라벨 확인
    for sh in s7.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t == "대출약정":
                checks.append(("대출약정 화살표", t, True))
            if t == "담보신탁계약":
                checks.append(("담보신탁계약 화살표", t, True))
            if "담보신탁 우선수익권" in t:
                checks.append(("담보신탁 우선수익권 화살표", t, True))

    # 5. 삭제된 라벨 확인
    deleted_texts = ["사업 시행", "자산관리", "책임준공", "Equity", "사모사채", "Bridge Loan"]
    remaining = []
    for sh in s7.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            for dt in deleted_texts:
                if dt in t:
                    remaining.append(t)
    checks.append(("삭제 라벨 잔존 없음", str(remaining[:3]), len(remaining) == 0))

    for name, detail, passed in checks:
        icon = "✓" if passed else "✗"
        print(f"  {icon} {name}: {detail[:50]}")

    pass_n = sum(1 for _, _, p in checks if p)
    fail_n = sum(1 for _, _, p in checks if not p)
    print(f"\n  결과: {pass_n} PASS / {fail_n} FAIL")

    print(f"\n{'=' * 70}")
    print(f"완료 — {OUTPUT}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()
