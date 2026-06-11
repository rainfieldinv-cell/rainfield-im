"""읽기 전용 검증 — 수정된 다단계 추출기로 천안 PDF 표 재추출 결과."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pdfplumber

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)
from modules import content_parser as CP

PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")
FOCUS = {8, 9, 7, 19, 20, 15}   # 상세 출력 페이지

with pdfplumber.open(PDF) as pdf:
    print(f"총 {len(pdf.pages)}페이지\n")
    print("=" * 92)
    print(f"{'pg':>3} | {'채택표':>5} | 표별 (행x열 / 전략 / 유효셀 / 통과)")
    print("=" * 92)
    summary = []
    detail = {}
    for pi, page in enumerate(pdf.pages):
        pno = pi + 1
        cands = CP._extract_tables_dual_strategy(page)
        lines = []
        accepted = []
        for ext, bbox, strat in cands:
            proc = CP._postprocess_table(ext)
            if not proc:
                continue
            valid = CP._is_valid_table(proc)
            r = len(proc); c = max((len(x) for x in proc), default=0)
            filled = CP._count_filled_cells(proc)
            lines.append(f"{r}x{c}/{strat}/{filled}/{'O' if valid else 'X'}")
            if valid:
                accepted.append(proc)
        if lines:
            print(f"{pno:>3} | {len(accepted):>5} | " + "  ".join(lines))
        if pno in FOCUS:
            detail[pno] = accepted

    print("\n" + "=" * 92)
    print("상세 출력 (채택된 표 셀 내용)")
    print("=" * 92)
    names = {8: "8p 금융조건표", 9: "9p 계좌/재무", 7: "7p Cash-in/out",
             19: "19p 차주 주주표", 20: "20p 재무제표", 15: "15p 분양사례"}
    for pno in [8, 9, 7, 19, 20, 15]:
        print(f"\n----- {names.get(pno, str(pno))} : 채택 {len(detail.get(pno, []))}개 -----")
        for ti, tbl in enumerate(detail.get(pno, [])):
            print(f"  [표{ti}] {len(tbl)}행×{max(len(r) for r in tbl)}열")
            for row in tbl[:8]:
                print("     ", [str(c or '').strip()[:16] for c in row])
