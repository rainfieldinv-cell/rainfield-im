"""고정 페이지 생성기+빌더 단독 검증 — 대전 PDF로 사모사채개요/투자구조도/ExecSummary 생성."""
import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT); os.chdir(ROOT)
from modules import page_builders
ov = os.environ.get("LAYOUT_OVERRIDE", "").strip()
if ov: page_builders.LAYOUT_PPTX_PATH = ov
import fitz
from modules.content_parser import parse_document
from modules.llm_structure import enrich_and_number
from modules import ai_slide_builders as A
from modules.page_builders import create_presentation_from_template, finalize_presentation
from modules.ppt_generator import save_presentation

PDF = os.path.join(ROOT, "[신영증권] 대전중구 서남부터미널 토지담보대출_IM_v3.0.pdf")
doc = fitz.open(PDF)
full = "\n".join(doc[i].get_text("text") for i in range(min(6, len(doc))))   # 앞부분
doc.close()

print("=== generate_sasae_overview ===")
r1 = A.generate_sasae_overview(full)
print("ok=", r1.get("ok"), json.dumps(r1.get("data"), ensure_ascii=False)[:600])
print("\n=== generate_executive_summary ===")
r2 = A.generate_executive_summary(full)
print("ok=", r2.get("ok"), json.dumps(r2.get("data"), ensure_ascii=False)[:600])
print("\n=== generate_investment_structure ===")
r3 = A.generate_investment_structure(full)
print("ok=", r3.get("ok"), json.dumps(r3.get("data"), ensure_ascii=False)[:600])

# 미니 덱: 표지 없이 3개 고정 슬라이드만
prs = create_presentation_from_template()
n = len(prs.slides)
if r2.get("ok"): A.build_slide_2_executive_summary(prs, r2["data"], business_name="대전 서남부터미널 토지담보대출", page_num=2)
if r1.get("ok"): A.build_slide_5_sasae_overview(prs, r1["data"], business_name="대전 서남부터미널 토지담보대출", page_num=5)
if r3.get("ok"): A.build_slide_7_investment_structure(prs, r3["data"], business_name="대전 서남부터미널 토지담보대출", page_num=7)
finalize_presentation(prs, n)
OUT = os.path.join(ROOT, "test_fixed_slides.pptx")
save_presentation(prs, OUT)
print("\n저장:", OUT, f"({len(prs.slides)}슬라이드)")
