"""읽기 전용 — 페이지 레이아웃 '자동 판별기' 시제품.
   각 페이지의 구성요소(텍스트/표/이미지)를 보고 유형(A/B/C/E/텍스트)을 규칙으로 자동 분류.
   변환 중(로딩) 이 판별을 자동 수행하는 게 목표. 26페이지 전체에 돌려 결과 확인."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import fitz
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")


def _inside(i, o, tol=2):
    return i[0] >= o[0]-tol and i[1] >= o[1]-tol and i[2] <= o[2]+tol and i[3] <= o[3]+tol


def extract_components(page):
    W, H = page.rect.width, page.rect.height
    tables, tbl_boxes = [], []
    for t in page.find_tables().tables:
        ext = t.extract() or []
        bbox = tuple(t.bbox); tbl_boxes.append(bbox)
        tables.append({"bbox": bbox, "rows": len(ext),
                       "cols": max((len(r) for r in ext), default=0), "ext": ext})
    images, texts = [], []
    for im in page.get_images(full=True):
        xref = im[0]
        rects = page.get_image_rects(xref)
        if not rects:
            continue
        r = rects[0]
        w_pt, h_pt = r[2]-r[0], r[3]-r[1]
        # 로고/아이콘 같은 작은 이미지 제외 (면적 작거나 상단 헤더영역)
        if w_pt * h_pt < 2500 or (r[1] < 60 and w_pt < 200):
            continue
        images.append({"bbox": tuple(r), "w": w_pt, "h": h_pt})
    for b in page.get_text("dict").get("blocks", []):
        if b.get("type") != 0:
            continue
        bbox = tuple(b["bbox"])
        if any(_inside(bbox, tb) for tb in tbl_boxes):
            continue
        txt = "".join(s["text"] for ln in b.get("lines", []) for s in ln.get("spans", [])).strip()
        if len(txt) >= 2 and "페이지" not in txt:
            texts.append({"bbox": bbox, "text": txt})
    return {"W": W, "H": H, "tables": tables, "images": images, "texts": texts}


def _is_label_value(tbl):
    """정규화 시 2열(구분|내용) 형태인지 — 각 행 비어있지않은 셀 2개 이하가 다수면 label-value."""
    rows = tbl["ext"]
    if not rows:
        return False
    le2 = sum(1 for r in rows if len([c for c in r if str(c or '').strip() and str(c or '').strip() != '•']) <= 2)
    return le2 >= max(2, len(rows) * 0.6)


def classify(comp):
    """구성요소 → 레이아웃 유형 자동 판별."""
    nt, ni = len(comp["tables"]), len(comp["images"])
    big_imgs = [im for im in comp["images"] if im["w"] * im["h"] > 8000]
    # 표 중 가장 큰 것
    main = max(comp["tables"], key=lambda t: t["rows"] * t["cols"], default=None)
    if nt == 0 and ni == 0:
        return "텍스트", "표·이미지 없음"
    if nt == 0 and ni >= 1:
        return "이미지위주", f"이미지{ni} 표0"
    if ni >= 1 and nt >= 1:
        if main and _is_label_value(main):
            return "E(데이터표+이미지표)", f"표{nt}(label-value {main['rows']}r) 이미지{ni}"
        return "B/C(이미지+표 좌우)", f"표{nt}({main['cols']}열) 이미지{ni}"
    # 표만
    if main and main["cols"] >= 5:
        return "A(전체폭 표)", f"표{nt} 최대{main['rows']}r×{main['cols']}c(넓음)"
    return "A/단일표", f"표{nt} 최대{main['rows']}r×{main['cols']}c"


doc = fitz.open(PDF)
print(f"{'pg':>3} | {'판별 유형':<22} | 근거")
print("-" * 78)
from collections import Counter
tally = Counter()
for pi in range(len(doc)):
    comp = extract_components(doc[pi])
    typ, why = classify(comp)
    tally[typ.split("(")[0]] += 1
    print(f"{pi+1:>3} | {typ:<22} | {why}")
doc.close()
print("\n유형별 집계:", dict(tally))
