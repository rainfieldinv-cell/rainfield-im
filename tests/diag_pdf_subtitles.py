import sys, io, os, re
import fitz
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
doc = fitz.open(os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf"))
txt = "\n".join(p.get_text() for p in doc)
doc.close()
for kw in ["인허가", "실시계획", "건축심의", "만기", "6개월", "Bridge", "1,640", "1640", "토지", "분양", "시공"]:
    idxs = [m.start() for m in re.finditer(kw, txt)]
    print(f"--- '{kw}' {len(idxs)}건 ---")
    for i in idxs[:2]:
        print("   ", txt[max(0, i-25):i+45].replace("\n", " "))
