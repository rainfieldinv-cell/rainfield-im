# -*- coding: utf-8 -*-
"""원본 PDF 페이지 텍스트 + 표 추출(영역별)로 원본 표 구조 확인."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import fitz

ROOT = r"C:\Users\jbzle\OneDrive\Desktop\종합\자동화\rainfield-im"
PDFS = {
    "대전": "[신영증권] 대전중구 서남부터미널 토지담보대출_IM_v3.0.pdf",
    "천안": "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf",
    "넷마블": "신영증권_구로_넷마블_지타워_담보대출_및_1종_수익증권_IM_v3_2.pdf",
}
key = sys.argv[1] if len(sys.argv) > 1 else "대전"
pages_arg = sys.argv[2] if len(sys.argv) > 2 else ""  # "3,4,5"
doc = fitz.open(os.path.join(ROOT, PDFS[key]))
print(f"=== {key} PDF : {len(doc)} pages ===\n")
want = [int(x) for x in pages_arg.split(",") if x.strip()] if pages_arg else range(1, len(doc)+1)
for pno in want:
    if pno < 1 or pno > len(doc):
        continue
    page = doc[pno-1]
    txt = page.get_text("text")
    print(f"────────── PDF p{pno} ──────────")
    print(txt[:2200])
    print()
