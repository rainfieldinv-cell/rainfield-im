"""
diag_slide7_fonts.py — STEP 4 전 폰트 검증. 빌더 실행 후 슬라이드 7 텍스트 shape 폰트 덤프.
파일 저장 안 함. 캐시 HIT 사용.
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
prs = create_presentation_from_template()
s7 = build_slide_7_investment_structure(prs, r7["data"],
                                        business_name="천안 부성2지구 도시개발사업", page_num=8)


def rinfo(r):
    sz = r.font.size.pt if r.font.size is not None else None
    try:
        col = str(r.font.color.rgb) if r.font.color and r.font.color.type is not None else "-"
    except Exception:
        col = "-"
    return f"name={r.font.name!r} size={sz} bold={r.font.bold} color={col}"


print(f"{'='*100}")
print("슬라이드 7 텍스트 shape 폰트 적용 결과")
print(f"{'='*100}")
for i, sh in enumerate(s7.shapes):
    if not sh.has_text_frame:
        continue
    txt = sh.text_frame.text.strip()
    if not txt:
        continue
    runs = [r for p in sh.text_frame.paragraphs for r in p.runs if r.text.strip()]
    if not runs:
        print(f"[{i:02d}] '{txt[:24]}' → (run 없음)")
        continue
    head = runs[0]
    print(f"[{i:02d}] '{txt[:24]}' → {rinfo(head)}")
    # 여러 run 의 폰트가 섞였는지 확인
    names = {r.font.name for r in runs}
    if len(names) > 1:
        print(f"      ⚠ 혼합 폰트: {names}")
print(f"\n{'='*100}\n진단 완료.")
