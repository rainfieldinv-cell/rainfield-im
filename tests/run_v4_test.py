"""
run_v4_test.py — test_ai_builders_v4.pptx 생성 + 슬라이드 7 전수 검증
빌더 코드(ai_slide_builders.py) 수정 없음. 생성 + 검증 전용.
"""
import sys, os, io, traceback

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
OUTPUT   = os.path.join(ROOT, "test_ai_builders_v4.pptx")
BIZ_NAME = "천안 부성2지구 도시개발사업"


def main():
    print("=" * 80)
    print("STEP 3: test_ai_builders_v4.pptx 생성")
    print("=" * 80)

    import fitz
    doc = fitz.open(PDF_PATH)
    pdf_text = "\n".join(page.get_text() for page in doc)
    doc.close()
    print(f"PDF 추출: {len(pdf_text):,} 자")

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

    print("\n--- API 호출 결과 ---")
    for label, r in [("S2", r2), ("S5", r5), ("S7", r7)]:
        cached = "캐시HIT" if r.get("cached") else "API호출"
        ok = "OK" if r.get("ok") else "FAIL"
        usage = r.get("usage", {})
        print(f"  {label}: {ok} {cached} usage={usage} {r.get('error','')}")

    if not all(r.get("ok") for r in [r2, r5, r7]):
        print("[FAIL] API 호출 실패")
        sys.exit(1)

    prs = create_presentation_from_template()
    n = len(prs.slides)
    print(f"\n템플릿 슬라이드: {n}")

    s2 = build_slide_2_executive_summary(prs, r2["data"], business_name=BIZ_NAME, page_num=2)
    s5 = build_slide_5_sasae_overview(prs, r5["data"], business_name=BIZ_NAME, page_num=6)
    s7 = build_slide_7_investment_structure(prs, r7["data"], business_name=BIZ_NAME, page_num=8)

    finalize_presentation(prs, n)
    save_presentation(prs, OUTPUT)
    print(f"\n저장 완료: {OUTPUT}")
    print(f"총 슬라이드 수: {len(prs.slides)}")

    # ── 슬라이드 7 shape 전수 조사 ──
    print(f"\n{'='*100}")
    print(f"슬라이드 7 shape 전수 조사 (총 {len(s7.shapes)}개)")
    print(f"{'='*100}")
    print(f"{'Idx':>3} | {'ShapeType':<14} | {'Name':<22} | {'L':>6} {'T':>6} | Text(30)")
    print(f"{'-'*3}-+-{'-'*14}-+-{'-'*22}-+-{'-'*6}-{'-'*6}-+-{'-'*32}")

    for i, sh in enumerate(s7.shapes):
        tp_raw = str(sh.shape_type)
        tp = tp_raw.split("(")[0].split(".")[-1] if "." in tp_raw else tp_raw
        name = (sh.name or "")[:22]
        l = sh.left / 360000 if sh.left else 0
        t = sh.top / 360000 if sh.top else 0
        txt = ""
        if sh.has_text_frame:
            txt = sh.text_frame.text.replace("\n", " | ")[:30]
        elif sh.shape_type == 19:
            tbl = sh.table
            cells = [tbl.cell(r, c).text.strip()[:12]
                     for r in range(len(tbl.rows)) for c in range(len(tbl.columns))
                     if tbl.cell(r, c).text.strip()]
            txt = (f"TBL {len(tbl.rows)}r{len(tbl.columns)}c: " +
                   ("/".join(cells) if cells else "(전부 비어있음)"))[:30]
        print(f"{i:3d} | {tp:<14} | {name:<22} | {l:6.2f} {t:6.2f} | {txt}")

    # ── 삭제 검증 ──
    print(f"\n{'='*100}")
    print("삭제 대상 잔존 여부 검증")
    print(f"{'='*100}")

    deleted_labels = {
        "[10] '사업 시행'": lambda t: t == "사업 시행" or t.strip() == "사업 시행",
        "[14] '자산관리'":  lambda t: t.strip() == "자산관리",
        "[22] 'Equity'":   lambda t: "Equity" in t,
        "[25] '책임준공'":  lambda t: "책임준공" in t,
    }
    texts = [sh.text_frame.text.strip() for sh in s7.shapes if sh.has_text_frame]
    for label, pred in deleted_labels.items():
        hits = [t for t in texts if pred(t)]
        icon = "OK(제거됨)" if not hits else "FAIL(잔존)"
        print(f"  {label}: {icon} {hits[:2] if hits else ''}")

    # TABLE 잔존 (자산관리자 2r, 투자자 7r)
    tables = [(i, sh.table) for i, sh in enumerate(s7.shapes) if sh.shape_type == 19]
    print(f"\n  남은 TABLE 개수: {len(tables)}")
    for i, tbl in tables:
        first = tbl.cell(0, 0).text.strip()
        print(f"    [{i}] {len(tbl.rows)}r x {len(tbl.columns)}c  첫셀='{first}'")

    has_7r = any(len(t.rows) == 7 for _, t in tables)
    print(f"  [17] 투자자 7r TABLE 잔존: {'FAIL(잔존)' if has_7r else 'OK(제거됨)'}")

    # PICTURE 잔존 (사업지 이미지)
    pics = [i for i, sh in enumerate(s7.shapes) if str(sh.shape_type).split("(")[0].endswith("PICTURE") or sh.shape_type == 13]
    print(f"  [27] PICTURE(사업지 이미지) 잔존: {'FAIL(잔존)' if pics else 'OK(제거됨)'} idx={pics}")

    # 푸터 [29]
    footer = [sh.text_frame.text.strip() for sh in s7.shapes
              if sh.has_text_frame and (sh.top or 0)/360000 > 18 and (sh.left or 0)/360000 < 1.0]
    print(f"  [29] 좌하단 사업명 푸터 잔존: {'FAIL(잔존)' if footer else 'OK(제거됨)'} {footer[:1]}")

    # ── PDF 원본 유지 요소 확인 ──
    print(f"\n{'='*100}")
    print("PDF 4번 사진 — 유지되어야 할 요소 확인")
    print(f"{'='*100}")
    keep = {
        "신탁사(신한자산신탁)": lambda: any("신탁" in t for t in texts) or any("신한" in (tbl.cell(r,c).text) for _,tbl in tables for r in range(len(tbl.rows)) for c in range(len(tbl.columns))),
    }
    # 박스(TABLE) 카운트 + 화살표 라벨
    arrow_labels = [t for t in texts if any(k in t for k in ["대출약정", "담보신탁", "공사도급", "우선수익", "도급계약"])]
    print(f"  남은 박스(TABLE) 수: {len(tables)}")
    print(f"  화살표/계약 라벨: {arrow_labels}")
    box_texts = []
    for i, tbl in tables:
        joined = " ".join(tbl.cell(r, c).text.strip()
                           for r in range(len(tbl.rows)) for c in range(len(tbl.columns))
                           if tbl.cell(r, c).text.strip())
        box_texts.append(f"[{i}] {joined}")
    print("  박스 내용:")
    for b in box_texts:
        print(f"    {b}")

    print(f"\n{'='*80}")
    print(f"완료 — {OUTPUT}")
    print(f"{'='*80}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
