"""
diag_slide7_step1.py — STEP 1 베이스 검증. 빌더 실행 후 슬라이드 7 남은 shape 덤프.
파일 저장 안 함 (v5 저장은 STEP 4). 캐시 HIT 사용.
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

import fitz
doc = fitz.open(os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf"))
pdf_text = "\n".join(p.get_text() for p in doc)
doc.close()

from modules.ai_slide_builders import (
    generate_investment_structure, build_slide_7_investment_structure,
)
from modules.page_builders import create_presentation_from_template

r7 = generate_investment_structure(pdf_text)
print(f"API: ok={r7.get('ok')} cached={r7.get('cached')}")

prs = create_presentation_from_template()
s7 = build_slide_7_investment_structure(prs, r7["data"],
                                        business_name="천안 부성2지구 도시개발사업", page_num=8)

print(f"\n{'='*100}")
print(f"STEP 1 베이스 — 슬라이드 7 남은 shape (총 {len(s7.shapes)}개)")
print(f"{'='*100}")
print(f"{'Idx':>3} | {'ShapeType':<16} | {'Name':<24} | {'L':>6} {'T':>6} {'W':>6} {'H':>6} | Text")
print(f"{'-'*3}-+-{'-'*16}-+-{'-'*24}-+-{'-'*6}-{'-'*6}-{'-'*6}-{'-'*6}-+-{'-'*30}")
for i, sh in enumerate(s7.shapes):
    tp = str(sh.shape_type).split("(")[0].split(".")[-1]
    name = (sh.name or "")[:24]
    l = sh.left/360000 if sh.left is not None else 0
    t = sh.top/360000 if sh.top is not None else 0
    w = sh.width/360000 if sh.width is not None else 0
    h = sh.height/360000 if sh.height is not None else 0
    txt = sh.text_frame.text.replace("\n", " | ")[:40] if sh.has_text_frame else ""
    print(f"{i:3d} | {tp:<16} | {name:<24} | {l:6.2f} {t:6.2f} {w:6.2f} {h:6.2f} | {txt}")
print(f"\n{'='*100}\n진단 완료.")
