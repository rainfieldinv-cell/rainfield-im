"""읽기 전용 — 8p/20p 후보별 내용 + _is_valid_table 탈락 사유."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pdfplumber
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)
from modules import content_parser as CP


def reject_reason(t):
    rows = len(t); cols = max((len(r) for r in t), default=0)
    if rows < 2: return "행<2"
    if cols < 2: return "열<2"
    if cols > 14: return "열>14"
    BUL = ('▶', '•', '■', '▪', '◆')
    for row in t:
        for c in row:
            if any(b in str(c or '') for b in BUL): return "불릿문자"
    total = sum(len(r) for r in t)
    empty = sum(1 for r in t for c in r if not str(c or '').strip())
    if total and empty/total > 0.95: return f"빈셀{empty/total*100:.0f}%>95"
    flat = [str(c).strip() for r in t for c in r if str(c or '').strip()]
    if len(set(flat)) < 2: return "유니크<2"
    longc = sum(1 for r in t for c in r if c and len(str(c)) > 80)
    if longc > 8: return f"긴셀{longc}>8"
    vh = sum(1 for c in t[0] if c and 0 < len(str(c).strip()) <= 40)
    if vh < 1: return "헤더유효셀<1"
    return "통과"


PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
with pdfplumber.open(PDF) as pdf:
    for pno in (8, 20):
        page = pdf.pages[pno - 1]
        print(f"\n{'='*80}\n[{pno}p] 후보 전체\n{'='*80}")
        for ext, bbox, strat in CP._extract_tables_dual_strategy(page):
            proc = CP._postprocess_table(ext)
            if not proc:
                print(f"  전략={strat}: postprocess 후 빈 표")
                continue
            r = len(proc); c = max(len(x) for x in proc)
            print(f"\n  전략={strat} {r}행×{c}열 filled={CP._count_filled_cells(proc)} "
                  f"→ {reject_reason(proc)}")
            for row in proc[:10]:
                print("       ", [str(x or '').strip()[:14] for x in row])
