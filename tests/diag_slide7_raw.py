"""
diag_slide7_raw.py — 삭제 전 원본 템플릿(슬라이드 8) shape 전수 덤프.
빌더 코드 수정 없음. 진단 전용. clone_slide_layout 직후 상태를 그대로 출력.
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

print(f"{'='*110}")
print(f"슬라이드 8(투자구조도) 원본 템플릿 — clone 직후 (총 {len(slide.shapes)}개)")
print(f"{'='*110}")
print(f"{'Idx':>3} | {'ShapeType':<16} | {'Name':<24} | {'L':>6} {'T':>6} {'W':>6} {'H':>6} | Text/내용")
print(f"{'-'*3}-+-{'-'*16}-+-{'-'*24}-+-{'-'*6}-{'-'*6}-{'-'*6}-{'-'*6}-+-{'-'*30}")

for i, sh in enumerate(slide.shapes):
    tp_raw = str(sh.shape_type)
    tp = tp_raw.split("(")[0].split(".")[-1] if "." in tp_raw else tp_raw
    name = (sh.name or "")[:24]
    l = sh.left / 360000 if sh.left is not None else 0
    t = sh.top / 360000 if sh.top is not None else 0
    w = sh.width / 360000 if sh.width is not None else 0
    h = sh.height / 360000 if sh.height is not None else 0
    txt = ""
    if sh.has_text_frame:
        txt = sh.text_frame.text.replace("\n", " | ")[:45]
    elif sh.shape_type == 19:
        tbl = sh.table
        cells = [tbl.cell(r, c).text.strip()[:12]
                 for r in range(len(tbl.rows)) for c in range(len(tbl.columns))
                 if tbl.cell(r, c).text.strip()]
        txt = (f"TBL {len(tbl.rows)}r{len(tbl.columns)}c: " +
               ("/".join(cells) if cells else "(비어있음)"))[:45]
    print(f"{i:3d} | {tp:<16} | {name:<24} | {l:6.2f} {t:6.2f} {w:6.2f} {h:6.2f} | {txt}")

print(f"\n{'='*110}")
print("진단 완료. 코드 수정 없음.")
