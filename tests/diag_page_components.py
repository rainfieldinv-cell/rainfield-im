"""읽기 전용 — 페이지 '구성요소 명세' 추출 시제품.
한 페이지에서 [텍스트블록 / 표 / 이미지]를 위치(bbox)·표 행열수와 함께 순서대로 뽑는다.
이게 '틀 먼저' 단계의 입력(명세)이 된다. 기존 코드 수정 없음."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import fitz   # PyMuPDF

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")


def _inside(inner, outer, tol=2):
    return (inner[0] >= outer[0]-tol and inner[1] >= outer[1]-tol and
            inner[2] <= outer[2]+tol and inner[3] <= outer[3]+tol)


def extract_components(page):
    """페이지 → 순서있는 구성요소 리스트.
    각 요소: {kind, bbox, rel(상대좌표 0~1), rows, cols, preview}"""
    W, H = page.rect.width, page.rect.height
    comps = []

    # 1) 표 (fitz find_tables) — bbox + 행/열
    tbl_boxes = []
    try:
        for t in page.find_tables().tables:
            bbox = tuple(t.bbox)
            tbl_boxes.append(bbox)
            ext = t.extract() or []
            rows = len(ext)
            cols = max((len(r) for r in ext), default=0)
            comps.append({"kind": "table", "bbox": bbox, "rows": rows, "cols": cols,
                          "preview": " / ".join(str(c or '') for c in (ext[0] if ext else []))[:40]})
    except Exception as e:
        print("  [표 추출 오류]", e)

    # 2) 텍스트/이미지 블록 (표 영역 안에 든 것은 제외)
    d = page.get_text("dict")
    for b in d.get("blocks", []):
        bbox = tuple(b["bbox"])
        if any(_inside(bbox, tb) for tb in tbl_boxes):
            continue   # 표 안의 텍스트 → 표가 이미 담당
        if b.get("type") == 1:   # 이미지
            comps.append({"kind": "image", "bbox": bbox, "rows": 0, "cols": 0,
                          "preview": f"{int(bbox[2]-bbox[0])}x{int(bbox[3]-bbox[1])}px"})
        else:                    # 텍스트
            txt = "".join(s["text"] for ln in b.get("lines", []) for s in ln.get("spans", [])).strip()
            if not txt:
                continue
            comps.append({"kind": "text", "bbox": bbox, "rows": 0, "cols": 0,
                          "preview": txt[:40]})

    # 위→아래, 좌→우 순 정렬
    comps.sort(key=lambda c: (round(c["bbox"][1] / 12), c["bbox"][0]))
    for c in comps:
        x0, y0, x1, y1 = c["bbox"]
        c["rel"] = (round(x0 / W, 2), round(y0 / H, 2), round(x1 / W, 2), round(y1 / H, 2))
    return comps, (W, H)


doc = fitz.open(PDF)
for pno in (3, 6, 8):       # 텍스트+표 섞인 대표 페이지
    page = doc[pno - 1]
    comps, (W, H) = extract_components(page)
    print(f"\n{'='*94}\n[{pno}p] 페이지크기 {W:.0f}x{H:.0f}pt, 구성요소 {len(comps)}개")
    print(f"{'순':>2} {'종류':<6} {'상대위치(x0,y0,x1,y1)':<26} {'표(행x열)':<9} 미리보기")
    print("-" * 94)
    for i, c in enumerate(comps):
        dim = f"{c['rows']}x{c['cols']}" if c["kind"] == "table" else ""
        print(f"{i:>2} {c['kind']:<6} {str(c['rel']):<26} {dim:<9} {c['preview']}")
doc.close()
