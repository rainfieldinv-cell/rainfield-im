"""LLM 페이지 구조화 단독 검증 — 대전 PDF 일부 페이지를 structure_page로 돌려 JSON 확인."""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)
import fitz
from modules.llm_structure import structure_page

PDF = os.path.join(ROOT, "[신영증권] 대전중구 서남부터미널 토지담보대출_IM_v3.0.pdf")
TARGET = [int(x) for x in sys.argv[1:]] or [7]   # 1-base 페이지

doc = fitz.open(PDF)
for pno in TARGET:
    raw = doc[pno - 1].get_text("text")
    print(f"\n{'='*90}\n[PDF p{pno}] 원문 {len(raw)}자")
    st = structure_page(raw, pno)
    if not st:
        print("  구조화 실패/None")
        continue
    print(f"  section_label: {st['section_label']!r}")
    print(f"  subtitle     : {st['subtitle']!r}")
    print(f"  intro        : {st['intro'][:90]!r}")
    print(f"  bullets({len(st['bullets'])}): {[b[:40] for b in st['bullets'][:4]]}")
    print(f"  source       : {st['source']!r}")
    print(f"  tables({len(st['tables'])}):")
    for t in st["tables"]:
        print(f"    · '{t.get('title','')}' kind={t.get('kind')} "
              f"header={t.get('header')} rows={len(t.get('rows',[]))}")
        for r in (t.get("rows") or [])[:3]:
            print(f"        {[str(c)[:16] for c in r]}")
    if st.get("_hallucinated"):
        print(f"  ⚠ 환각 의심 숫자: {st['_hallucinated']}")
doc.close()
