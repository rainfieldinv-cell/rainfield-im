"""읽기 전용 — 천안 PDF 표 추출 실제 실행. 원본 라인-감지 vs _is_valid_table 필터 결과 비교."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pdfplumber

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)
from modules import content_parser as CP
PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")

with pdfplumber.open(PDF) as pdf:
    print(f"총 {len(pdf.pages)}페이지\n")
    for pi, page in enumerate(pdf.pages):
        # 1) 원본 pdfplumber lines 전략 감지
        raw = page.find_tables(table_settings={"vertical_strategy": "lines",
                                               "horizontal_strategy": "lines"}) or []
        if not raw:
            continue
        print(f"===== page {pi+1}: lines전략 {len(raw)}개 표 감지 =====")
        for ti, t in enumerate(raw):
            ext = t.extract()
            nrows = len(ext)
            ncols = max((len(r) for r in ext), default=0)
            # 빈 셀 비율
            total = sum(len(r) for r in ext)
            empty = sum(1 for r in ext for c in r if not str(c or "").strip())
            er = empty / total * 100 if total else 0
            # 헤더 셀 수
            hdr = ext[0] if ext else []
            valid_hdr = sum(1 for c in hdr if c and 0 < len(str(c).strip()) <= 30)
            valid = CP._is_valid_table(ext)
            print(f"  [표{ti}] {nrows}r×{ncols}c 빈셀{er:.0f}% 유효헤더셀{valid_hdr} "
                  f"→ _is_valid_table={valid} {'(통과)' if valid else '(버려짐)'}")
            # 첫 3행 셀 내용 일부
            for r in ext[:3]:
                cells = [(str(c or '').strip()[:14]) for c in r]
                print(f"        {cells}")
        # _extract_page_tables 최종 결과(필터 통과분)
        res = CP._extract_page_tables(page)
        print(f"  → 최종 채택 표: {len(res['tables'])}개\n")
