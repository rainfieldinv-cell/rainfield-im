"""
텔레그램 알림 전송 (IM 변환기 전용)
--------------------------------------------------------------
PPT 가 실제로 생성/다운로드될 때, 중계 서버로 알림 요청을 보냅니다.
중계 서버가 그 요청을 받아 텔레그램 그룹방으로 메시지를 전달합니다.

[원칙]
· fire-and-forget — 백그라운드 스레드로 보내므로 PPT 생성/다운로드를
  절대 지연시키거나 방해하지 않습니다. 모든 오류는 로그만 남기고 무시합니다.
· 토큰은 코드에 박지 않습니다. 호출하는 쪽(app.py)에서 st.secrets 로 읽어
  이 함수에 넘겨줍니다. (브라우저에 노출되지 않음)

[의존성]
· 표준 라이브러리(urllib)만 사용 → requirements.txt 수정 불필요.
"""

import json
import logging
import threading
import urllib.request

log = logging.getLogger("rainfield-notify")

# ── 설정 (보통 수정할 일 없음) ─────────────────────────────
RELAY_URL = "https://rainfield-notify-relay.onrender.com/notify"
PROJECT_NAME = "IM 변환기"   # 대시보드에 표시되는 이름과 동일
# ───────────────────────────────────────────────────────────


def _post(token: str, detail: str) -> None:
    """실제 전송 (백그라운드 스레드에서 실행). st.* 를 절대 호출하지 않습니다."""
    try:
        body = json.dumps(
            {"project": PROJECT_NAME, "event": "download", "detail": detail or "IM"}
        ).encode("utf-8")
        req = urllib.request.Request(
            RELAY_URL,
            data=body,
            method="POST",
            headers={
                "Content-Type": "application/json",
                "X-Notify-Token": token or "",
            },
        )
        # Render 무료플랜이 잠들어 있으면 콜드스타트로 수십 초 걸릴 수 있어
        # 넉넉히 기다립니다. (백그라운드라 사용자 화면은 영향 없음)
        with urllib.request.urlopen(req, timeout=60) as resp:
            resp.read()
    except Exception as e:
        # 네트워크 실패, 401, 타임아웃 등 무엇이든 조용히 무시 (기능에 영향 없음)
        log.warning("텔레그램 알림 전송 실패(무시): %s", e)


def notify_download(token: str, detail: str = None) -> None:
    """
    다운로드 알림을 백그라운드로 보냅니다. (호출 즉시 반환)

    Parameters
    ----------
    token : str
        NOTIFY_TOKEN 값. app.py 에서 st.secrets 로 읽어 넘깁니다.
    detail : str, optional
        실제 PPT 파일명. 없으면 "IM" 으로 전송됩니다.
    """
    try:
        threading.Thread(target=_post, args=(token, detail), daemon=True).start()
    except Exception as e:
        log.warning("알림 스레드 시작 실패(무시): %s", e)
