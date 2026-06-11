"""읽기 전용 — 8p에 추출 가능한 이미지(위치도/조감도)가 있는지 점검."""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import fitz
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")

doc = fitz.open(PDF)
for pno in (8,):
    page = doc[pno - 1]
    W, H = page.rect.width, page.rect.height
    print(f"[{pno}p] 페이지 {W:.0f}x{H:.0f}")
    # 1) get_images (xref 기반 래스터 이미지)
    imgs = page.get_images(full=True)
    print(f"  get_images: {len(imgs)}개")
    for i, im in enumerate(imgs):
        xref = im[0]
        try:
            rects = page.get_image_rects(xref)
        except Exception:
            rects = []
        base = doc.extract_image(xref)
        print(f"    [{i}] xref={xref} {base.get('width')}x{base.get('height')}px ext={base.get('ext')} "
              f"위치={[tuple(round(v) for v in r) for r in rects]}")
    # 2) dict 블록 중 image 타입
    blocks = page.get_text("dict").get("blocks", [])
    img_blocks = [b for b in blocks if b.get("type") == 1]
    print(f"  dict image blocks: {len(img_blocks)}개")
    for b in img_blocks:
        bb = tuple(round(v) for v in b["bbox"])
        print(f"    bbox={bb} ({bb[2]-bb[0]}x{bb[3]-bb[1]}px)")
    # 3) 드로잉(벡터) 여부 — 위치도가 벡터일 수도
    drawings = page.get_drawings()
    print(f"  vector drawings: {len(drawings)}개 (벡터 도형/선)")
doc.close()
