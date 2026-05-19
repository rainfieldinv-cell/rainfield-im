"""
slide_mapping.py
─────────────────────────────────────────────────────────
PDF → PPT 변환 시 각 슬라이드를 어떤 방식으로 처리할지 정의합니다.

method 값:
  "python"     — 규칙 기반 처리 (템플릿 복제 + 텍스트/이미지 교체)
  "claude_api" — Claude API 호출로 내용 생성/변환
  "hybrid"     — python 처리 후 Claude API 보조

source 값:
  "template"     — 템플릿에서 복제 (PDF 내용 불필요)
  "pdf_direct"   — PDF 해당 페이지를 거의 그대로 변환
  "pdf_summary"  — PDF 전체 또는 일부를 요약·변환
  "pdf_image"    — PDF 이미지 추출 후 배치
─────────────────────────────────────────────────────────
"""

SLIDE_MAPPING = [
    {
        "slide_num": 1,
        "slide_name": "표지",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": True,
        "description": "사업명, 날짜, 표지 사진 (사용자 업로드)",
    },
    {
        "slide_num": 2,
        "slide_name": "Executive Summary",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": "all",
        "needs_ai": True,
        "needs_images": True,
        "description": "낮은 인허가 리스크 / 시공사 리스크 / 토지확보 리스크 / 높은 분양성 / 만기 — 5가지 핵심 요약",
    },
    {
        "slide_num": 3,
        "slide_name": "목차",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": False,
        "description": "5형식 자동 분할 (이미 구현됨)",
    },
    {
        "slide_num": 4,
        "slide_name": "섹션 1 divider — 사모사채 개요",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": True,
        "description": "섹션 구분 슬라이드, 원형 이미지 3장 슬롯",
    },
    {
        "slide_num": 5,
        "slide_name": "1.1 본건 사모사채 개요",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [3, 4],
        "needs_ai": True,
        "needs_images": False,
        "description": "Bridge Loan → 사모사채 관점 변환, 핵심 조건 재정리",
    },
    {
        "slide_num": 6,
        "slide_name": "섹션 2 divider — 금융개요",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": True,
        "description": "섹션 구분 슬라이드, 원형 이미지 3장 슬롯",
    },
    {
        "slide_num": 7,
        "slide_name": "2.1 투자구조도",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [6],
        "needs_ai": True,
        "needs_images": False,
        "description": "텍스트 → 다이어그램 데이터 변환 (관계자, 자금 흐름)",
    },
    {
        "slide_num": 8,
        "slide_name": "2.2 본건 기초자산 금융조건 1/3",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [5, 6, 7],
        "needs_ai": True,
        "needs_images": False,
        "description": "PDF p5~p7 핵심 조건 추리기 — 첫 번째 파트",
    },
    {
        "slide_num": 9,
        "slide_name": "2.2 본건 기초자산 금융조건 2/3",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [5, 6, 7],
        "needs_ai": True,
        "needs_images": False,
        "description": "PDF p5~p7 핵심 조건 추리기 — 두 번째 파트",
    },
    {
        "slide_num": 10,
        "slide_name": "2.2 본건 기초자산 금융조건 3/3",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [5, 6, 7],
        "needs_ai": True,
        "needs_images": False,
        "description": "PDF p5~p7 핵심 조건 추리기 — 세 번째 파트",
    },
    {
        "slide_num": 11,
        "slide_name": "섹션 3 divider — 본건 사업 개요",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": True,
        "description": "섹션 구분 슬라이드, 원형 이미지 3장 슬롯",
    },
    {
        "slide_num": 12,
        "slide_name": "3.1 사업개요",
        "source": "pdf_direct",
        "method": "python",
        "pdf_pages": [8],
        "needs_ai": False,
        "needs_images": True,
        "description": "PDF p8 표 그대로 + 사업지 이미지 1장",
    },
    {
        "slide_num": 13,
        "slide_name": "3.2 입지분석 지도",
        "source": "pdf_image",
        "method": "hybrid",
        "pdf_pages": [14],
        "needs_ai": True,
        "needs_images": True,
        "description": "PDF 이미지 활용, AI 보조 캡션 생성",
    },
    {
        "slide_num": 14,
        "slide_name": "3.2 입지분석 텍스트",
        "source": "pdf_direct",
        "method": "python",
        "pdf_pages": [14],
        "needs_ai": False,
        "needs_images": False,
        "description": "PDF p14 텍스트 그대로",
    },
    {
        "slide_num": 15,
        "slide_name": "3.3 사업수지",
        "source": "pdf_direct",
        "method": "python",
        "pdf_pages": [9],
        "needs_ai": False,
        "needs_images": False,
        "description": "PDF p9 사업수지 표만 추출",
    },
    {
        "slide_num": 16,
        "slide_name": "3.4 토지 확보 현황",
        "source": "pdf_direct",
        "method": "python",
        "pdf_pages": [10],
        "needs_ai": False,
        "needs_images": True,
        "description": "PDF p10 그대로 (표 + 지도 이미지)",
    },
    {
        "slide_num": 17,
        "slide_name": "3.5 토지 수용",
        "source": "pdf_direct",
        "method": "python",
        "pdf_pages": [11],
        "needs_ai": False,
        "needs_images": False,
        "description": "PDF p11 그대로 + 설명 한 줄",
    },
    {
        "slide_num": 18,
        "slide_name": "3.6 청약단지 시세비교",
        "source": "pdf_direct",
        "method": "python",
        "pdf_pages": [15],
        "needs_ai": False,
        "needs_images": True,
        "description": "PDF p15 표 + 이미지",
    },
    {
        "slide_num": 19,
        "slide_name": "3.6 최근 분양사례",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [16],
        "needs_ai": True,
        "needs_images": False,
        "description": "표 → 카드 4+1 재배치",
    },
    {
        "slide_num": 20,
        "slide_name": "3.6 기입주 사례",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [17],
        "needs_ai": True,
        "needs_images": False,
        "description": "표 → 카드 6+1 재배치",
    },
    {
        "slide_num": 21,
        "slide_name": "섹션 4 divider — Appendix",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": True,
        "description": "섹션 구분 슬라이드, 원형 이미지 3장 슬롯",
    },
    {
        "slide_num": 22,
        "slide_name": "4.1 차주 개요",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [19, 20],
        "needs_ai": True,
        "needs_images": False,
        "description": "PDF p19+p20 통합 요약",
    },
    {
        "slide_num": 23,
        "slide_name": "4.2 본PF 주요조건",
        "source": "pdf_summary",
        "method": "claude_api",
        "pdf_pages": [21],
        "needs_ai": True,
        "needs_images": False,
        "description": "PDF p21 핵심 조건 추리기",
    },
    {
        "slide_num": 24,
        "slide_name": "4.3 인허가 관련 공문",
        "source": "pdf_image",
        "method": "python",
        "pdf_pages": [22, 23],
        "needs_ai": False,
        "needs_images": True,
        "description": "PDF p22+p23 공문 이미지 통합 배치",
    },
    {
        "slide_num": 25,
        "slide_name": "4.4 시공사 양해각서",
        "source": "pdf_image",
        "method": "python",
        "pdf_pages": [24],
        "needs_ai": False,
        "needs_images": True,
        "description": "PDF p24 이미지 그대로 배치",
    },
    {
        "slide_num": 26,
        "slide_name": "연락처",
        "source": "template",
        "method": "python",
        "pdf_pages": [],
        "needs_ai": False,
        "needs_images": False,
        "description": "Rainfield 연락처 (이미 구현됨)",
    },
]


# ─────────────────────────────────────────────
# 유틸리티 함수
# ─────────────────────────────────────────────

_VALID_METHODS  = {"python", "claude_api", "hybrid"}
_VALID_SOURCES  = {"template", "pdf_direct", "pdf_summary", "pdf_image"}


def get_slides_needing_ai() -> list:
    """Claude API 호출이 필요한 슬라이드 번호 리스트를 반환합니다."""
    return [s["slide_num"] for s in SLIDE_MAPPING if s.get("needs_ai")]


def get_slides_python_only() -> list:
    """Python 규칙 기반만으로 처리 가능한 슬라이드 번호 리스트를 반환합니다."""
    return [s["slide_num"] for s in SLIDE_MAPPING if s["method"] == "python"]


def get_pdf_pages_for_slide(slide_num: int) -> list:
    """특정 슬라이드가 참조하는 PDF 페이지 번호 리스트를 반환합니다."""
    for s in SLIDE_MAPPING:
        if s["slide_num"] == slide_num:
            pages = s["pdf_pages"]
            if pages == "all":
                return list(range(1, 27))   # 천안 PDF 기준 1~26
            return list(pages)
    return []


def validate_mapping() -> dict:
    """
    SLIDE_MAPPING 테이블을 검증합니다.

    Returns
    -------
    {"errors": [...], "warnings": [...]}
    """
    errors   = []
    warnings = []
    nums     = [s["slide_num"] for s in SLIDE_MAPPING]

    # 중복 slide_num
    seen = set()
    for n in nums:
        if n in seen:
            errors.append(f"slide_num={n} 중복")
        seen.add(n)

    # 1~26 연속성 확인
    expected = set(range(1, len(SLIDE_MAPPING) + 1))
    missing  = expected - set(nums)
    if missing:
        errors.append(f"누락된 slide_num: {sorted(missing)}")

    # method / source 유효성
    for s in SLIDE_MAPPING:
        if s["method"] not in _VALID_METHODS:
            errors.append(f"slide_num={s['slide_num']}: 유효하지 않은 method={s['method']!r}")
        if s["source"] not in _VALID_SOURCES:
            errors.append(f"slide_num={s['slide_num']}: 유효하지 않은 source={s['source']!r}")

    # needs_ai=True 인데 method='python' 이면 경고
    for s in SLIDE_MAPPING:
        if s.get("needs_ai") and s["method"] == "python":
            warnings.append(f"slide_num={s['slide_num']}: needs_ai=True 이지만 method='python'")

    return {"errors": errors, "warnings": warnings}
