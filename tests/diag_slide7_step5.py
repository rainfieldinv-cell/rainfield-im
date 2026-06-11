"""
diag_slide7_step5.py — 방식 A(표 기반) 빌드 검증. 저장 안 함. 캐시 HIT.
LAYOUT_OVERRIDE 환경변수로 템플릿 경로 교체(잠금 회피).
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

from modules import page_builders
from modules.ai_slide_builders import (
    generate_investment_structure, build_slide_7_investment_structure,
)
from modules.page_builders import create_presentation_from_template

ov = os.environ.get("LAYOUT_OVERRIDE", "").strip()
if ov:
    page_builders.LAYOUT_PPTX_PATH = ov

r7 = generate_investment_structure(pdf_text)
prs = create_presentation_from_template()
s7 = build_slide_7_investment_structure(prs, r7["data"],
                                        business_name="천안 부성2지구 도시개발사업", page_num=8)

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def hexfill(cell):
    try:
        f = cell.fill
        if f.type is not None and f.fore_color is not None and f.fore_color.type is not None:
            return str(f.fore_color.rgb)
    except Exception:
        pass
    return "-"


print(f"\n{'='*100}")
print(f"STEP 5-2 빌드 — 슬라이드 7 전체 shape (총 {len(s7.shapes)}개)")
print(f"{'='*100}")
for i, sh in enumerate(s7.shapes):
    tp = str(sh.shape_type).split("(")[0].split(".")[-1].strip()
    l = sh.left/360000 if sh.left is not None else 0
    t = sh.top/360000 if sh.top is not None else 0
    w = sh.width/360000 if sh.width is not None else 0
    h = sh.height/360000 if sh.height is not None else 0
    li = l/2.54; ti = t/2.54; wi = w/2.54; hi = h/2.54   # 인치 변환
    if sh.has_table:
        tb = sh.table
        print(f"[{i:02d}] TABLE {len(tb.rows)}r  L={li:.2f}\" T={ti:.2f}\" W={wi:.2f}\" H={hi:.2f}\"")
        for r in range(len(tb.rows)):
            c = tb.cell(r, 0)
            print(f"        row{r}: 배경={hexfill(c):<8} '{c.text.strip()[:20]}'")
    else:
        txt = sh.text_frame.text.replace("\n", " ")[:26] if sh.has_text_frame else ""
        nm = "Conn" if "Connector" in (sh.name or "") else ""
        print(f"[{i:02d}] {tp:<11}{nm:<5} L={li:5.2f}\" T={ti:5.2f}\" W={wi:5.2f}\" H={hi:5.2f}\" {txt}")

# 요약
tables = [sh for sh in s7.shapes if sh.has_table]
conns = [sh for sh in s7.shapes if "Connector" in (sh.name or "")]
labels = [sh.text_frame.text.strip() for sh in s7.shapes
          if sh.has_text_frame and sh.text_frame.text.strip() in
          {"신탁계약", "담보신탁 우선수익권", "공사도급계약", "대출약정"}]
print(f"\n요약: 표 {len(tables)}개 / 연결선 {len(conns)}개 / 라벨 {len(labels)}개 {sorted(labels)}")
daeju = [sh for sh in tables if sh.table.cell(0, 0).text.strip() == "대주"]
if daeju:
    print(f"대주 표 행수: {len(daeju[0].table.rows)}r (Tr.C 삭제 확인 = {'OK' if len(daeju[0].table.rows)==3 else 'FAIL'})")
