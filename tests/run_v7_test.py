"""
run_v7_test.py — test_ai_builders_v7.pptx 생성 + 슬라이드 1·7 검증.
수정1(슬7 헤더색) + 수정2(슬1 레이아웃). 캐시 HIT. LAYOUT_OVERRIDE 로 잠금 회피.
"""
import sys, os, io, traceback
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
OUTPUT   = os.path.join(ROOT, "test_ai_builders_v7.pptx")
BIZ_NAME = "천안 부성2지구 도시개발사업"


def hexfill(cell):
    try:
        f = cell.fill
        if f.type is not None and f.fore_color is not None and f.fore_color.type is not None:
            return str(f.fore_color.rgb)
    except Exception:
        pass
    return "-"


def cm(v):
    return (v or 0) / 360000


def main():
    print("=" * 80)
    print("STEP 5-4: test_ai_builders_v7.pptx 생성")
    print("=" * 80)

    import fitz
    doc = fitz.open(PDF_PATH)
    pdf_text = "\n".join(p.get_text() for p in doc)
    doc.close()

    from modules import page_builders
    from modules.ai_slide_builders import (
        generate_executive_summary, generate_sasae_overview, generate_investment_structure,
        build_slide_2_executive_summary, build_slide_5_sasae_overview,
        build_slide_7_investment_structure,
    )
    from modules.page_builders import create_presentation_from_template, finalize_presentation
    from modules.ppt_generator import save_presentation

    ov = os.environ.get("LAYOUT_OVERRIDE", "").strip()
    if ov:
        page_builders.LAYOUT_PPTX_PATH = ov
        print(f"[override] {ov}")

    r2 = generate_executive_summary(pdf_text)
    r5 = generate_sasae_overview(pdf_text)
    r7 = generate_investment_structure(pdf_text)
    for lb, r in [("S2", r2), ("S5", r5), ("S7", r7)]:
        print(f"  {lb}: {'OK' if r.get('ok') else 'FAIL'} {'캐시HIT' if r.get('cached') else 'API'}")
    if not all(r.get("ok") for r in [r2, r5, r7]):
        sys.exit(1)

    prs = create_presentation_from_template()
    n = len(prs.slides)
    s2 = build_slide_2_executive_summary(prs, r2["data"], business_name=BIZ_NAME, page_num=2)
    build_slide_5_sasae_overview(prs, r5["data"], business_name=BIZ_NAME, page_num=6)
    s7 = build_slide_7_investment_structure(prs, r7["data"], business_name=BIZ_NAME, page_num=8)
    finalize_presentation(prs, n)
    save_presentation(prs, OUTPUT)
    print(f"\n저장: {OUTPUT}")

    # ── 슬라이드 7: 헤더색 ──
    print(f"\n{'='*90}\n[슬라이드 7] 박스 헤더색 (수정1)\n{'='*90}")
    exp7 = {"신탁사": "0070C0", "시공사": "2E75B6", "차주": "013A73", "대주": "8C4A59"}
    got7 = {}
    for sh in s7.shapes:
        if sh.has_table:
            h = sh.table.cell(0, 0).text.strip()
            got7[h] = hexfill(sh.table.cell(0, 0))
    for role, c in exp7.items():
        ok = got7.get(role) == c
        print(f"  [{'OK ' if ok else 'FAIL'}] {role}: {got7.get(role)} (기대 {c})")
    print(f"  차주=013A73 고정 확인: {'OK' if got7.get('차주')=='013A73' else 'FAIL'}")

    # ── 슬라이드 1: 섹션/육각형 좌표 ──
    print(f"\n{'='*90}\n[슬라이드 1] 섹션 <> 육각형 도형 좌표 (수정2)\n{'='*90}")
    print(f"{'섹션':<6} {'shape':<12} {'L':>6} {'T':>6} {'W':>6} {'H':>6}")
    hexes = []
    for sh in s2.shapes:
        if "육각형" in (sh.name or ""):
            hexes.append(sh)
    # top 으로 섹션 묶기
    hexes.sort(key=lambda s: cm(s.top))
    for sh in hexes:
        kind = "바깥" if cm(sh.width) >= 20.5 else "안쪽"
        print(f"  {kind:<6} {'육각형':<10} {cm(sh.left):6.2f} {cm(sh.top):6.2f} {cm(sh.width):6.2f} {cm(sh.height):6.2f}")

    # 육각형 폭 통일 검증
    outers = [sh for sh in hexes if cm(sh.width) >= 20.5]
    inners = [sh for sh in hexes if cm(sh.width) < 20.5]
    out_w = {round(cm(s.width), 2) for s in outers}
    out_l = {round(cm(s.left), 2) for s in outers}
    in_w = {round(cm(s.width), 2) for s in inners}
    in_l = {round(cm(s.left), 2) for s in inners}
    print(f"\n  바깥 육각형 L={out_l} W={out_w}  → {'통일 OK' if len(out_l)==1 and len(out_w)==1 else 'FAIL'}")
    print(f"  안쪽 육각형 L={in_l} W={in_w}  → {'통일 OK' if len(in_l)==1 and len(in_w)==1 else 'FAIL'}")

    # 섹션 헤더(GROUP) top 으로 간격 확인
    groups = sorted([cm(sh.top) for sh in s2.shapes if str(sh.shape_type).split('(')[0].split('.')[-1].strip() == "GROUP"])
    print(f"\n  섹션 헤더(GROUP) top: {[round(g,2) for g in groups]}")
    if len(groups) >= 3:
        print(f"  헤더 간격: {round(groups[1]-groups[0],2)} / {round(groups[2]-groups[1],2)} cm")

    print(f"\n완료 — {OUTPUT}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc(); sys.exit(1)
