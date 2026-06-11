"""통합 검증 — 실제 파이프라인(parse → build_full_presentation)에 틀-우선 라우터가
   배선됐는지 확인. 천안 PDF를 파싱해 페이지별 유형 판별 + 전체 PPT 생성."""
import sys, io, os, zipfile
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from modules import page_builders
ov = os.environ.get("LAYOUT_OVERRIDE", "").strip()
if ov:
    page_builders.LAYOUT_PPTX_PATH = ov

from modules.content_parser import parse_document
from modules.frame_builders import classify_page
from modules.page_builders import build_full_presentation

_PDFS = {
    "천안": ("★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf", "변환결과_천안_v2.pptx", "천안 부성2지구 도시개발사업"),
    "대전": ("[신영증권] 대전중구 서남부터미널 토지담보대출_IM_v3.0.pdf", "변환결과_대전_v98.pptx", "대전 서남부터미널 토지담보대출"),
}
_key = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] in _PDFS else "천안"
_pdf, _out, _biz = _PDFS[_key]
PDF = os.path.join(ROOT, _pdf)
OUT = os.path.join(ROOT, _out)

pages = parse_document(PDF)
print(f"파싱: {len(pages)}페이지")

# LLM 켜져 있으면 페이지 구조화 + 섹션/소제목 번호 선처리 + Investment Highlights 생성
exec_summary_data = None
if os.environ.get("RAINFIELD_LLM", "").strip() in ("1", "true", "True"):
    from modules.llm_structure import enrich_and_number
    print("LLM 구조화 + 섹션 번호 부여 중...")
    enrich_and_number(pages, debug=True, pdf_path=PDF)
    # Investment Highlights: ★원본 'Executive Summary' 페이지만 보고 생성(전체 문서 요약 금지)
    import fitz
    from modules.ai_slide_builders import generate_executive_summary
    _doc = fitz.open(PDF)
    _es_pages = [d.get_text("text") for d in _doc
                 if "Executive Summary" in (d.get_text("text") or "")]
    if not _es_pages:   # 폴백: 앞 2페이지
        _es_pages = [_doc[i].get_text("text") for i in range(min(2, len(_doc)))]
    _front = "\n".join(_es_pages)
    _doc.close()
    _es = generate_executive_summary(_front)
    if _es.get("ok"):
        exec_summary_data = _es["data"]
        print("Investment Highlights 생성 완료")

ppt = build_full_presentation(
    business_name=_biz,
    year="2026", month_en="June",
    pages=pages,
    exec_summary_data=exec_summary_data,
)
with open(OUT, "wb") as f:
    f.write(ppt)

dup = [nm for nm, c in Counter(zipfile.ZipFile(OUT).namelist()).items() if c > 1]
from pptx import Presentation
prs = Presentation(io.BytesIO(ppt))
print(f"\n저장: {OUT}  ({len(prs.slides)}슬라이드)  중복파트: {'있음' if dup else '없음(정상)'}")
