"""읽기 전용 — 파서가 페이지별로 뱉은 결과(제목/부제목/본문/표) 진단."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)
from modules.content_parser import parse_document

PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
pages = parse_document(PDF)
TARGET = [int(x) for x in sys.argv[1:]] or list(range(len(pages)))

for i in TARGET:
    if i >= len(pages):
        continue
    p = pages[i]
    tbls = p.get("tables") or []
    print(f"\n{'='*90}\n[페이지 {i+1}]")
    print(f"  section_title: {p.get('section_title','')[:70]!r}")
    print(f"  subtitle     : {p.get('subtitle','')[:70]!r}")
    bt = p.get('body_text','')
    print(f"  body_text({len(bt)}자): {bt[:120]!r}")
    twt = p.get('text_without_tables','')
    print(f"  text_without_tables({len(twt)}자): {twt[:120]!r}")
    print(f"  표 {len(tbls)}개:")
    for ti, t in enumerate(tbls):
        cols = max((len(r) for r in t), default=0)
        print(f"    [표{ti+1}] {len(t)}행 × {cols}열")
        for r in t[:3]:
            print(f"        {[str(c or '')[:12] for c in r]}")
