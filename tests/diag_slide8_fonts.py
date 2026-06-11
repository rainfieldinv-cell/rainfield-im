"""
diag_slide8_fonts.py — 레이아웃 슬라이드 8 원본 박스/라벨의 실제 폰트 덤프.
빌더 수정 없음. clone 직후 raw 상태의 run.font.name/size/color 출력.
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from modules.page_builders import create_presentation_from_template, clone_slide_layout

prs = create_presentation_from_template()
slide = clone_slide_layout(prs, "investment_structure")


def run_info(run):
    nm = run.font.name
    sz = run.font.size.pt if run.font.size is not None else None
    bold = run.font.bold
    try:
        col = str(run.font.color.rgb) if run.font.color and run.font.color.type is not None else "-"
    except Exception:
        col = "-"
    return f"name={nm!r} size={sz} bold={bold} color={col}"


print(f"{'='*110}")
print("슬라이드 8 원본 — 텍스트가 있는 shape 의 run 폰트")
print(f"{'='*110}")
for i, sh in enumerate(slide.shapes):
    l = sh.left/360000 if sh.left is not None else 0
    t = sh.top/360000 if sh.top is not None else 0
    if sh.has_text_frame:
        txt = sh.text_frame.text.strip()
        if not txt:
            continue
        print(f"\n[{i:02d}] TEXT_BOX L={l:.2f} T={t:.2f}  text='{txt[:30]}'")
        for pi, p in enumerate(sh.text_frame.paragraphs):
            for ri, r in enumerate(p.runs):
                if r.text.strip():
                    print(f"      p{pi}r{ri} '{r.text[:20]}' → {run_info(r)}")
    elif sh.shape_type == 19:  # TABLE
        tbl = sh.table
        print(f"\n[{i:02d}] TABLE {len(tbl.rows)}r{len(tbl.columns)}c L={l:.2f} T={t:.2f}")
        for ri in range(len(tbl.rows)):
            for ci in range(len(tbl.columns)):
                cell = tbl.cell(ri, ci)
                ctxt = cell.text.strip()
                if not ctxt:
                    continue
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text.strip():
                            print(f"      cell[{ri},{ci}] '{r.text[:20]}' → {run_info(r)}")
print(f"\n{'='*110}\n진단 완료.")
