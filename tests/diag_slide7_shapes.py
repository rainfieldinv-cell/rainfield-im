"""
diag_slide7_shapes.py — 슬라이드 7 빌더 실행 후 shape 전수 조사
코드 수정 없음. 진단 전용.
"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")

import fitz
doc = fitz.open(PDF_PATH)
pdf_text = "\n".join(page.get_text() for page in doc)
doc.close()

from modules.ai_slide_builders import (
    generate_investment_structure,
    build_slide_7_investment_structure,
)
from modules.page_builders import create_presentation_from_template, finalize_presentation

# Claude API (캐시)
r7 = generate_investment_structure(pdf_text)
if not r7.get("ok"):
    print(f"[FAIL] API: {r7.get('error')}")
    sys.exit(1)
print(f"API OK (cached={r7.get('cached')})")

# 빌더 실행
prs = create_presentation_from_template()
n_template = len(prs.slides)
s7 = build_slide_7_investment_structure(prs, r7["data"], business_name="천안 부성2지구 도시개발사업", page_num=8)

print(f"\n{'='*100}")
print(f"슬라이드 7 shape 전수 조사 (총 {len(s7.shapes)}개)")
print(f"{'='*100}")
print(f"{'Idx':>3} | {'ShapeType':<20} | {'Name':<25} | {'L(cm)':>7} {'T(cm)':>7} {'W(cm)':>7} {'H(cm)':>7} | Text (30자)")
print(f"{'-'*3}-+-{'-'*20}-+-{'-'*25}-+-{'-'*7}-{'-'*7}-{'-'*7}-{'-'*7}-+-{'-'*40}")

for i, sh in enumerate(s7.shapes):
    tp_raw = str(sh.shape_type)
    tp = tp_raw.split("(")[0].split(".")[-1] if "." in tp_raw else tp_raw
    name = sh.name[:25] if sh.name else ""
    l = sh.left / 360000 if sh.left else 0
    t = sh.top / 360000 if sh.top else 0
    w = sh.width / 360000 if sh.width else 0
    h = sh.height / 360000 if sh.height else 0

    txt = ""
    if sh.has_text_frame:
        txt = sh.text_frame.text.replace("\n", " | ")[:40]
    elif sh.shape_type == 19:  # TABLE
        tbl = sh.table
        cells = []
        for r in range(len(tbl.rows)):
            for c in range(len(tbl.columns)):
                ct = tbl.cell(r, c).text.strip()[:15]
                if ct:
                    cells.append(ct)
        txt = f"TBL {len(tbl.rows)}r{len(tbl.columns)}c: " + "/".join(cells) if cells else f"TBL {len(tbl.rows)}r{len(tbl.columns)}c: (전부 비어있음)"

    print(f"{i:3d} | {tp:<20} | {name:<25} | {l:7.2f} {t:7.2f} {w:7.2f} {h:7.2f} | {txt}")

# 분류
print(f"\n{'='*100}")
print("PDF 원본 요소 매칭 확인")
print(f"{'='*100}")

pdf_elements = {
    "신탁사 박스": None,
    "차주 박스": None,
    "시공사 박스": None,
    "대주(Tr) 박스": None,
    "담보신탁계약 라벨": None,
    "담보신탁 우선수익권 라벨": None,
    "공사도급계약 라벨": None,
    "대출약정 라벨": None,
}

delete_candidates = []

for i, sh in enumerate(s7.shapes):
    txt = ""
    if sh.has_text_frame:
        txt = sh.text_frame.text.strip()
    elif sh.shape_type == 19:
        tbl = sh.table
        all_cells = []
        for r in range(len(tbl.rows)):
            for c in range(len(tbl.columns)):
                all_cells.append(tbl.cell(r, c).text.strip())
        txt = " ".join(all_cells)

    # PDF 원본 매칭
    if sh.shape_type == 19 and hasattr(sh, 'table'):
        tbl = sh.table
        r0 = tbl.cell(0, 0).text.strip() if len(tbl.rows) > 0 else ""
        if "신탁" in r0:
            pdf_elements["신탁사 박스"] = f"[{i}] {r0}/{tbl.cell(1,0).text.strip() if len(tbl.rows)>1 else ''}"
        elif "차주" in r0 or "시행" in r0:
            pdf_elements["차주 박스"] = f"[{i}] {r0}/{tbl.cell(1,0).text.strip() if len(tbl.rows)>1 else ''}"
        elif "시공" in r0:
            pdf_elements["시공사 박스"] = f"[{i}] {r0}/{tbl.cell(1,0).text.strip() if len(tbl.rows)>1 else ''}"
        elif len(tbl.rows) == 4:
            pdf_elements["대주(Tr) 박스"] = f"[{i}] {len(tbl.rows)}r - {' / '.join(c.strip() for c in [tbl.cell(r,0).text for r in range(len(tbl.rows))])}"

    if sh.has_text_frame:
        if "담보신탁계약" in txt and "우선수익권" not in txt:
            pdf_elements["담보신탁계약 라벨"] = f"[{i}] {txt[:30]}"
        elif "우선수익권" in txt:
            pdf_elements["담보신탁 우선수익권 라벨"] = f"[{i}] {txt[:30]}"
        elif "공사" in txt and "도급" in txt:
            pdf_elements["공사도급계약 라벨"] = f"[{i}] {txt[:30]}"
        elif "대출약정" in txt:
            pdf_elements["대출약정 라벨"] = f"[{i}] {txt[:30]}"

    # 삭제 대상 판별
    is_pdf_element = False
    if sh.shape_type == 19 and hasattr(sh, 'table'):
        tbl = sh.table
        r0 = tbl.cell(0, 0).text.strip() if len(tbl.rows) > 0 else ""
        if any(k in r0 for k in ["신탁", "차주", "시행", "시공"]):
            is_pdf_element = True
        elif len(tbl.rows) == 4:  # 대주 박스
            is_pdf_element = True

    if sh.has_text_frame:
        t = sh.text_frame.text.strip()
        if t in ["담보신탁계약", "담보신탁 우선수익권", "대출약정"] or ("공사" in t and "도급" in t):
            is_pdf_element = True
        # 헤더/타이틀 영역도 유지
        if "금융 개요" in t or "투자구조도" in t:
            is_pdf_element = True
        # 인트로 텍스트 유지
        l = sh.left / 360000 if sh.left else 0
        tt = sh.top / 360000 if sh.top else 0
        if abs(l - 1.09) < 0.3 and abs(tt - 2.29) < 0.3:
            is_pdf_element = True
        # 페이지 번호 유지
        if abs(l - 18.53) < 0.3 and tt > 18:
            is_pdf_element = True

    # LINE shapes — PDF에 있는 화살표선인지 판별
    tp_raw = str(sh.shape_type)
    if "LINE" in tp_raw or "FREEFORM" in tp_raw:
        # 헤더/푸터 구분선은 유지
        t_cm = sh.top / 360000 if sh.top else 0
        if abs(t_cm - 3.87) < 0.3 or abs(t_cm - 18.16) < 0.3:
            is_pdf_element = True
        # 나머지 LINE은 구조도 화살표로 간주 → 유지/삭제 개별 판단 필요
        else:
            is_pdf_element = True  # 우선 유지로 분류

    # PLACEHOLDER (헤더) 유지
    if "PLACEHOLDER" in tp_raw:
        is_pdf_element = True

    # PICTURE — PDF에 없음 (사업지 이미지)
    if "PICTURE" in tp_raw:
        is_pdf_element = False

    if not is_pdf_element:
        reason = ""
        if sh.shape_type == 19:
            tbl = sh.table
            all_empty = all(tbl.cell(r, 0).text.strip() == "" for r in range(len(tbl.rows)))
            if all_empty:
                reason = "빈 TABLE"
            elif len(tbl.rows) == 7:
                reason = "투자자 7r TABLE"
            elif len(tbl.rows) == 2:
                reason = "자산관리자/SPC TABLE"
            else:
                reason = "불명 TABLE"
        elif sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t == "":
                reason = "빈 TEXT_BOX"
            else:
                reason = f"불필요 라벨: '{t[:20]}'"
        elif "PICTURE" in tp_raw:
            reason = "사업지 이미지"
        else:
            reason = "불명"
        delete_candidates.append((i, reason))

print("\n── PDF 원본 요소 매칭 ──")
for name, val in pdf_elements.items():
    icon = "✓" if val else "✗"
    print(f"  {icon} {name}: {val or '못 찾음'}")

print(f"\n── 삭제 대상 ({len(delete_candidates)}개) ──")
for idx, reason in delete_candidates:
    sh = s7.shapes[idx]
    tp_raw = str(sh.shape_type).split(".")[-1].split("(")[0]
    print(f"  shape[{idx:02d}] {tp_raw:<15} → {reason}")

print(f"\n{'='*100}")
print("진단 완료. 코드 수정 없음.")
