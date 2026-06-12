"""7단계 '전체 미리보기' — 생성 PPT를 슬라이드별 이미지로 변환(LibreOffice → PDF → PNG).

★읽기 전용: 본문 생성/검수 로직과 무관. 변환 실패 시 예외를 잡아 (이미지없음, 에러메시지)를
  돌려주어 앱이 절대 죽지 않게 한다(Streamlit Cloud=리눅스 기준).

서버 의존성(packages.txt): libreoffice, fonts-nanum(한글), poppler-utils(pdf2image용).
파이썬 의존성(requirements.txt): pdf2image.
"""
import io
import os
import shutil
import tempfile
import subprocess


def _find_soffice():
    """LibreOffice 실행 파일 경로 탐색(리눅스 우선, 로컬 대비 폴백)."""
    for cand in ("soffice", "libreoffice",
                 "/usr/bin/soffice", "/usr/bin/libreoffice",
                 "/usr/lib/libreoffice/program/soffice"):
        p = shutil.which(cand) if os.sep not in cand else (cand if os.path.exists(cand) else None)
        if p:
            return p
    return None


def ppt_to_images(ppt_bytes: bytes, max_pages=None, dpi: int = 120):
    """PPT bytes → (images: list[bytes(PNG)], error: str|None).

    max_pages: 앞 N페이지만 변환(테스트 모드). None이면 전체.
    실패하면 ([], '사람이 읽을 수 있는 원인') 을 반환 — 호출부에서 메시지만 표시하면 됨.
    """
    if not ppt_bytes:
        return [], "변환할 PPT가 없습니다. 먼저 4단계에서 PPT를 생성하세요."

    soffice = _find_soffice()
    if not soffice:
        return [], ("LibreOffice(soffice)가 설치되어 있지 않습니다. "
                    "packages.txt에 'libreoffice' 추가 후 재배포가 필요합니다.")

    tmp = tempfile.mkdtemp(prefix="rf_preview_")
    try:
        pptx_path = os.path.join(tmp, "deck.pptx")
        with open(pptx_path, "wb") as f:
            f.write(ppt_bytes)

        # 1) PPT → PDF (LibreOffice headless). HOME을 임시폴더로 줘야 프로필 생성 충돌이 없음.
        env = dict(os.environ)
        env["HOME"] = tmp
        try:
            proc = subprocess.run(
                [soffice, "--headless", "--norestore", "--nofirststartwizard",
                 "--convert-to", "pdf", "--outdir", tmp, pptx_path],
                capture_output=True, timeout=240, env=env,
            )
        except subprocess.TimeoutExpired:
            return [], "LibreOffice 변환 시간 초과(240초). 페이지 수가 많으면 '앞 3페이지만'으로 먼저 시도하세요."
        except Exception as e:
            return [], f"LibreOffice 실행 실패: {e}"

        pdf_path = os.path.join(tmp, "deck.pdf")
        if not os.path.exists(pdf_path):
            err = (proc.stderr or b"").decode("utf-8", "ignore").strip()
            return [], f"PDF 변환 실패: {err[:300] or 'soffice가 PDF를 생성하지 못했습니다.'}"

        # 2) PDF → PNG (pdf2image + poppler)
        try:
            from pdf2image import convert_from_path
        except Exception as e:
            return [], f"pdf2image 임포트 실패(requirements.txt 확인): {e}"

        last = max_pages if (isinstance(max_pages, int) and max_pages > 0) else None
        try:
            pages = convert_from_path(pdf_path, dpi=dpi, first_page=1, last_page=last)
        except Exception as e:
            return [], f"PDF→이미지 변환 실패(poppler-utils 필요): {e}"

        images = []
        for img in pages:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            images.append(buf.getvalue())

        if not images:
            return [], "변환된 이미지가 없습니다(빈 PDF)."
        return images, None

    except Exception as e:                                # 어떤 예외든 앱은 살린다
        return [], f"미리보기 생성 중 오류: {e}"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
