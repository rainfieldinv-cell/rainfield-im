"""
run_v5_test.py — test_ai_builders_v5.pptx 생성 + 슬라이드 7 검증 (STEP 5-4-H 신규 방식).
빌더 코드 수정 없음. 생성 + 검증 전용. 캐시 HIT 사용.
"""
import sys, os, io, traceback
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
OUTPUT   = os.path.join(ROOT, "test_ai_builders_v5.pptx")
BIZ_NAME = "천안 부성2지구 도시개발사업"


def main():
    print("=" * 80)
    print("STEP 4: test_ai_builders_v5.pptx 생성")
    print("=" * 80)

    import fitz
    doc = fitz.open(PDF_PATH)
    pdf_text = "\n".join(p.get_text() for p in doc)
    doc.close()
    print(f"PDF 추출: {len(pdf_text):,} 자")

    from modules.ai_slide_builders import (
        generate_executive_summary, generate_sasae_overview,
        generate_investment_structure,
        build_slide_2_executive_summary, build_slide_5_sasae_overview,
        build_slide_7_investment_structure,
    )
    from modules import page_builders
    from modules.page_builders import create_presentation_from_template, finalize_presentation
    from modules.ppt_generator import save_presentation

    # OneDrive 플레이스홀더 + PowerPoint 잠금 우회: 로컬 임시 복사본으로 템플릿 경로 교체
    _override = os.environ.get("LAYOUT_OVERRIDE", "").strip()
    if _override:
        page_builders.LAYOUT_PPTX_PATH = _override
        print(f"[override] LAYOUT_PPTX_PATH → {_override}")

    r2 = generate_executive_summary(pdf_text)
    r5 = generate_sasae_overview(pdf_text)
    r7 = generate_investment_structure(pdf_text)

    print("\n--- API 호출 결과 ---")
    for label, r in [("S2", r2), ("S5", r5), ("S7", r7)]:
        cached = "캐시HIT" if r.get("cached") else "API호출"
        print(f"  {label}: {'OK' if r.get('ok') else 'FAIL'} {cached} usage={r.get('usage', {})}")

    if not all(r.get("ok") for r in [r2, r5, r7]):
        print("[FAIL] API 호출 실패"); sys.exit(1)

    prs = create_presentation_from_template()
    n = len(prs.slides)

    build_slide_2_executive_summary(prs, r2["data"], business_name=BIZ_NAME, page_num=2)
    build_slide_5_sasae_overview(prs, r5["data"], business_name=BIZ_NAME, page_num=6)
    s7 = build_slide_7_investment_structure(prs, r7["data"], business_name=BIZ_NAME, page_num=8)

    finalize_presentation(prs, n)
    save_presentation(prs, OUTPUT)
    print(f"\n저장 완료: {OUTPUT}")
    print(f"총 슬라이드 수: {len(prs.slides)}")

    # ── 슬라이드 7 전체 shape ──
    print(f"\n{'='*100}")
    print(f"슬라이드 7 전체 shape (총 {len(s7.shapes)}개)")
    print(f"{'='*100}")
    print(f"{'Idx':>3} | {'종류':<12} | {'L':>6} {'T':>6} {'W':>6} {'H':>6} | Text")
    boxes, arrows, labels = [], [], []
    LBL = {"신탁계약", "담보신탁 우선수익권", "공사도급계약", "대출약정"}
    BOX = {"신탁사", "시공사", "차주", "대주"}
    for i, sh in enumerate(s7.shapes):
        tp = str(sh.shape_type).split("(")[0].split(".")[-1]
        l = sh.left/360000 if sh.left is not None else 0
        t = sh.top/360000 if sh.top is not None else 0
        w = sh.width/360000 if sh.width is not None else 0
        h = sh.height/360000 if sh.height is not None else 0
        txt = sh.text_frame.text.replace("\n", " ")[:30] if sh.has_text_frame else ""
        print(f"{i:3d} | {tp:<12} | {l:6.2f} {t:6.2f} {w:6.2f} {h:6.2f} | {txt}")
        if tp == "AUTO_SHAPE":
            boxes.append((i, txt))
        if tp in ("LINE",) and not sh.has_text_frame and i >= 6:
            arrows.append(i)
        if sh.has_text_frame and txt.strip() in LBL:
            labels.append(txt.strip())

    # 연결선(Connector)만 카운트: name 에 Connector 포함
    connectors = [i for i, sh in enumerate(s7.shapes)
                  if "Connector" in (sh.name or "")]

    # ── PDF p4 시각 일치 검증 ──
    all_text = [sh.text_frame.text.strip() for sh in s7.shapes if sh.has_text_frame]
    joined = " ".join(all_text)

    box_roles = [t for _, t in boxes if t.strip() in BOX]
    box_companies = [sh.text_frame.text.strip() for sh in s7.shapes
                     if str(sh.shape_type).split("(")[0].endswith("AUTO_SHAPE")
                     and sh.text_frame.text.strip() in
                     ("신한자산신탁", "포스코이앤씨", "더함도시개발", "Tr.A", "Tr.B")]
    found_labels = [t for t in LBL if t in all_text]

    # 폰트 확인
    fonts = set()
    for sh in s7.shapes:
        if str(sh.shape_type).split("(")[0].endswith("AUTO_SHAPE") or \
           (sh.has_text_frame and sh.text_frame.text.strip() in LBL):
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.text.strip() and r.font.name:
                        fonts.add(r.font.name)

    # 금지 요소
    has_amount = "억원" in joined or "억" in joined
    has_asset_mgr = "자산관리자" in joined or "자산관리" in joined
    has_investor = any(k in joined for k in ["투자자", "Equity", "삼성증권", "KT Estate", "우미건설"])
    pic_count = sum(1 for sh in s7.shapes
                    if str(sh.shape_type).split("(")[0].endswith("PICTURE") or sh.shape_type == 13)

    print(f"\n{'='*100}")
    print("PDF 천안 부성2지구 p4 구조도 시각 일치 검증")
    print(f"{'='*100}")
    checks = [
        ("박스 4개(신탁사/시공사/차주/대주)", "있음", f"{len(set(box_roles))}개 {sorted(set(box_roles))}", len(set(box_roles)) == 4),
        ("  └ 회사명 5개(신한/포스코/더함/Tr.A/Tr.B)", "있음", f"{len(set(box_companies))}개", len(set(box_companies)) == 5),
        ("화살표(연결선) 3개", "있음", f"{len(connectors)}개", len(connectors) == 3),
        ("라벨 4개", "있음", f"{len(found_labels)}개 {sorted(found_labels)}", len(found_labels) == 4),
        ("폰트: 피플폰트 Bold", "적용", f"{sorted(fonts)}", fonts == {"피플폰트 Bold"}),
        ("금액 표기 없음", "없음", "억/억원 미검출" if not has_amount else "검출됨!", not has_amount),
        ("자산관리자 박스 없음", "없음", "미검출" if not has_asset_mgr else "검출됨!", not has_asset_mgr),
        ("사업명 이미지(PICTURE) 없음", "없음", f"PICTURE {pic_count}개", pic_count == 0),
        ("우측 투자자 박스 없음", "없음", "미검출" if not has_investor else "검출됨!", not has_investor),
    ]
    for name, pdf, actual, ok in checks:
        print(f"  [{'OK' if ok else 'FAIL'}] {name:<40} | PDF={pdf:<5} | v5={actual}")

    pass_n = sum(1 for *_, ok in checks if ok)
    print(f"\n  결과: {pass_n}/{len(checks)} PASS")
    print(f"\n{'='*80}\n완료 — {OUTPUT}\n{'='*80}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc(); sys.exit(1)
