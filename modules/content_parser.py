"""
content_parser.py
─────────────────────────────────────────────────────────
PDF / Word 문서에서 슬라이드별 내용(제목·부제목·본문)을 추출합니다.

반환 형식:
  [
    {"section_title": "01  사모사채개요",
     "subtitle":      "1.1 현재 자금조달 구조",
     "body_text":     "...본문..."},
    ...
  ]
─────────────────────────────────────────────────────────
"""

import io
import os
import re
from typing import List, Dict

PageData = Dict[str, str]

# ─────────────────────────────────────────────
# [DEBUG 플래그]
# True 로 바꾸면 각 페이지 파싱 결과가 콘솔에 출력됩니다.
# 배포 시에는 False 로 유지하세요.
# ─────────────────────────────────────────────
DEBUG_PARSER = True


# ─────────────────────────────────────────────
# [확실한 페이지번호·헤더 패턴 — 화이트리스트 원칙]
# 아래 패턴에 완전히 일치하는 줄만 제거합니다.
# 의심스러우면 남깁니다.
# ─────────────────────────────────────────────
_OBVIOUS_FOOTER_RES = [
    re.compile(r'^[-–—]\s*\d{1,3}\s*[-–—]$'),   # "- 3 -"
    re.compile(r'^\d{1,3}\s*/\s*\d{1,3}$'),       # "3 / 25"
    re.compile(r'^페이지\s*\d+$'),                  # "페이지 3"
    re.compile(r'^Page\s+\d+$', re.IGNORECASE),   # "Page 3"
    re.compile(r'^\d{1,3}$'),                      # 단독 숫자 "3"
]

# 목차 페이지 감지 — 이 문자열이 페이지 텍스트(소문자)에 포함되면 목차로 판단
_TOC_MARKERS = ['목차', 'c o n t e n t s', 'contents']


# ──────────────────────────────────────────────────────
# 공개 API
# ──────────────────────────────────────────────────────

def parse_document(file_path: str) -> List[PageData]:
    """
    PDF 또는 Word(.docx) 파일을 읽어 슬라이드 단위로 분리한 목록을 반환합니다.

    Parameters
    ----------
    file_path : 로컬 파일 경로 (.pdf 또는 .docx)

    Returns
    -------
    list[dict] — 각 항목이 하나의 content 슬라이드에 대응:
        section_title  제목 텍스트 박스에 들어갈 내용
        subtitle       부제목 텍스트 박스에 들어갈 내용
        body_text      본문 텍스트 박스에 들어갈 내용
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        with open(file_path, "rb") as f:
            return _parse_docx_bytes(f.read())
    elif ext == ".pdf":
        with open(file_path, "rb") as f:
            return map_pdf_pages_to_slides(f.read(), debug=DEBUG_PARSER)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext!r}  (.pdf 또는 .docx만 지원)")


def parse_document_from_bytes(data: bytes, filename: str) -> List[PageData]:
    """
    Streamlit 파일 업로더에서 받은 bytes를 바로 파싱합니다.

    Parameters
    ----------
    data     : 업로드된 파일의 bytes
    filename : 원본 파일명 (확장자 판별에 사용)
    """
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".docx":
        return _parse_docx_bytes(data)
    elif ext == ".pdf":
        return map_pdf_pages_to_slides(data, debug=DEBUG_PARSER)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext!r}  (.pdf 또는 .docx만 지원)")


# ──────────────────────────────────────────────────────
# [페이지 매핑 — 핵심 공개 함수]
# ──────────────────────────────────────────────────────

def map_pdf_pages_to_slides(data: bytes, debug: bool = False) -> List[PageData]:
    """
    PDF 각 페이지를 분석해 PPT content 슬라이드 데이터 목록을 반환합니다.
    1:1 매핑 원칙: 각 PDF 콘텐츠 페이지 → 1개 content 슬라이드 데이터.

    자동 제외 규칙:
      - 1페이지: 표지 (PPT는 별도 커버 슬라이드를 사용)
      - 목차 포함 페이지 (목차 / CONTENTS 감지)
      - 텍스트 거의 없는 페이지 (15자 미만)

    섹션 제목 전파:
      - 현재 페이지 헤더가 너무 짧으면 (10자 미만) 이전 페이지의 section_title 유지

    Parameters
    ----------
    data  : PDF 파일 bytes
    debug : True 이면 각 페이지 처리 결과를 콘솔에 출력

    Returns
    -------
    list[PageData]  — build_full_presentation()에 바로 전달 가능한 형식
    """
    import fitz
    import pdfplumber

    fitz_doc     = fitz.open(stream=data, filetype="pdf")
    plumber_pdf  = pdfplumber.open(io.BytesIO(data))
    total        = fitz_doc.page_count
    result: List[PageData] = []
    last_section_title = ""

    if debug:
        print(f"\n[content_parser] PDF 총 {total}페이지 분석 시작")
        print("─" * 60)

    for i, pdf_page in enumerate(fitz_doc):
        page_num = i + 1
        full_text_lower = (pdf_page.get_text("text") or "").lower().strip()

        # ── 규칙 1: 1페이지는 표지로 간주, 항상 제외 ──────────
        if page_num == 1:
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 표지 (1페이지 자동 제외)")
            continue

        # ── 규칙 2: 목차 페이지 제외 ───────────────────────────
        if any(marker in full_text_lower for marker in _TOC_MARKERS):
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 목차 페이지")
            continue

        # ── 페이지 텍스트 추출 (fitz) ──────────────────────────
        pd = _extract_pdf_page(pdf_page, debug=debug)
        if pd is None:
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 텍스트 없음")
            continue

        # ── 규칙 3: 내용 부족 페이지 제외 ─────────────────────
        combined = (pd["section_title"] + pd["subtitle"] + pd["body_text"]).strip()
        if len(combined) < 15:
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 내용 부족 ({len(combined)}자)")
            continue

        # ── 표 추출 (pdfplumber) ───────────────────────────────
        table_data = _extract_page_tables(plumber_pdf.pages[i], debug=debug)
        pd.update(table_data)   # tables, table_bboxes, text_without_tables 추가

        # ── 섹션 제목 전파 ─────────────────────────────────────
        _SEC_NUM = re.compile(r'^0[1-9][\s\.\-]')
        has_sec_num = bool(_SEC_NUM.match(pd["section_title"].strip()))

        if has_sec_num:
            last_section_title = pd["section_title"]
        elif last_section_title:
            # {**pd, ...} 로 tables 등 기존 키 전부 보존
            pd = {
                **pd,
                "section_title": last_section_title,
                "subtitle":      pd["subtitle"] or pd["section_title"],
            }

        if debug:
            sec_preview  = pd["section_title"][:30]
            body_preview = pd["body_text"][:40].replace("\n", " ")
            tbl_count    = len(pd.get("tables", []))
            print(f"[p{page_num:02d}/{total}] KEEP  | sec={sec_preview!r:35} | "
                  f"body={len(pd['body_text'])}자 | tables={tbl_count}개 | {body_preview!r}")

        result.append(pd)

    plumber_pdf.close()
    fitz_doc.close()

    if debug:
        print("─" * 60)
        print(f"[content_parser] 최종 content 슬라이드 수: {len(result)}장\n")

    return result


# ──────────────────────────────────────────────────────
# [필터 함수 — 확실한 헤더/푸터만 제거]
# ──────────────────────────────────────────────────────

def clean_page_text(lines: List[str], debug: bool = False) -> List[str]:
    """
    페이지 텍스트 줄 목록에서 확실한 페이지번호·헤더 패턴만 제거합니다.
    화이트리스트 원칙: 의심스러우면 남깁니다.

    제거 대상:
      - "- 3 -", "3 / 25", "페이지 3", "Page 3", 단독 숫자

    Parameters
    ----------
    lines : 텍스트 줄 목록
    debug : True 이면 제거된 줄을 콘솔에 출력

    Returns
    -------
    필터링된 줄 목록
    """
    result = []
    for line in lines:
        stripped = line.strip()
        removed = any(p.fullmatch(stripped) for p in _OBVIOUS_FOOTER_RES)
        if removed:
            if debug:
                print(f"  [clean_page_text 제거] {stripped!r}")
        else:
            result.append(line)
    return result


# ──────────────────────────────────────────────────────
# Word (.docx) 파서
# ──────────────────────────────────────────────────────

def _parse_docx_bytes(data: bytes) -> List[PageData]:
    import docx

    doc = docx.Document(io.BytesIO(data))
    pages: List[PageData] = []

    current_section_title = ""
    current_subtitle      = ""
    body_lines: List[str] = []
    in_page = False

    def _flush():
        nonlocal current_subtitle, body_lines, in_page
        if in_page:
            pages.append({
                "section_title": current_section_title,
                "subtitle":      current_subtitle,
                "body_text":     "\n".join(body_lines).strip(),
            })
        current_subtitle = ""
        body_lines       = []
        in_page          = False

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style = para.style.name if para.style else ""

        if _is_heading_level(style, 1):
            _flush()
            current_section_title = text

        elif _is_heading_level(style, 2):
            _flush()
            current_subtitle = text
            in_page = True

        elif _is_heading_level(style, 3):
            _flush()
            current_subtitle = text
            in_page = True

        else:
            if not in_page:
                in_page = True
            body_lines.append(text)

    _flush()

    if not pages and current_section_title:
        pages.append({
            "section_title": "",
            "subtitle":      current_section_title,
            "body_text":     "\n".join(body_lines).strip(),
        })

    return pages


def _is_heading_level(style_name: str, level: int) -> bool:
    """'Heading 1', '제목 1', 'heading1' 등 다양한 스타일명을 허용합니다."""
    s = style_name.lower()
    if f"heading {level}" in s:
        return True
    if f"제목 {level}" in s:
        return True
    if f"heading{level}" in s:
        return True
    return False


# ──────────────────────────────────────────────────────
# PDF 파서 — 페이지 단위 텍스트 추출
# ──────────────────────────────────────────────────────

def _extract_pdf_page(page, debug: bool = False) -> "PageData | None":
    """
    PDF 한 페이지에서 헤더(상단 20%) 텍스트를 제목/부제목으로,
    나머지를 본문으로 추출합니다.
    본문에는 clean_page_text()로 명확한 페이지번호만 제거합니다.
    """
    page_h = page.rect.height
    header_thresh = page_h * 0.20

    lines: List[tuple] = []
    blocks = page.get_text("dict", flags=0)["blocks"]

    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            y0       = line["bbox"][1]
            txt      = ""
            max_size = 0.0
            bold     = False
            for span in line["spans"]:
                t = span["text"].strip()
                if t:
                    txt += t + " "
                    if span["size"] > max_size:
                        max_size = span["size"]
                    if span["flags"] & 16:
                        bold = True
            txt = txt.strip()
            if txt and len(txt) >= 2:
                lines.append((y0, txt, max_size, bold))

    if not lines:
        return None

    lines.sort(key=lambda x: x[0])

    header_lines = [(y, t, sz, b) for y, t, sz, b in lines if y < header_thresh]
    body_lines   = [(y, t, sz, b) for y, t, sz, b in lines if y >= header_thresh]

    # 헤더에서 제목 / 부제목 분리 (가장 큰 폰트 = 섹션 제목)
    section_title = ""
    subtitle      = ""
    if header_lines:
        max_sz = max(sz for _, _, sz, _ in header_lines)
        title_parts    = [t for _, t, sz, _ in header_lines if sz >= max_sz - 1.0]
        subtitle_parts = [t for _, t, sz, _ in header_lines if sz < max_sz - 1.0]
        section_title  = " ".join(title_parts).strip()
        subtitle       = " ".join(subtitle_parts).strip()

    # 본문 — 명확한 페이지번호 패턴만 제거 (화이트리스트 원칙)
    raw_body_texts = [t for _, t, _, _ in body_lines]
    clean_body     = clean_page_text(raw_body_texts, debug=debug)
    body_text      = "\n".join(clean_body).strip()

    return {
        "section_title": section_title,
        "subtitle":      subtitle,
        "body_text":     body_text,
    }


# ──────────────────────────────────────────────────────
# PDF 파서 — pdfplumber 표 추출
# ──────────────────────────────────────────────────────

def _postprocess_table(table: list) -> list:
    """
    추출된 표 데이터를 정제합니다.

    처리 순서:
      1. 각 셀의 None → "", \\n → 공백, 앞뒤 공백 제거
      2. 모든 셀이 빈 행 제거
      3. 행이 1개뿐인 표는 단순 제목 줄로 간주해 버림
      4. 동일 헤더가 반복되는 경우 첫 번째만 유지 (병합 셀 잔재 정리)
    """
    if not table:
        return []

    # 1. 셀 정제
    cleaned = [
        [str(cell).replace('\n', ' ').strip() if cell is not None else ""
         for cell in row]
        for row in table
    ]

    # 2. 모든 셀이 빈 행 제거
    cleaned = [row for row in cleaned if any(c for c in row)]

    # 2b. 빈 열 제거 — 90% 이상 비어있는 열은 phantom column으로 간주
    if cleaned:
        max_cols = max(len(r) for r in cleaned)
        keep = []
        for c in range(max_cols):
            vals = [r[c] if c < len(r) else "" for r in cleaned]
            non_empty = sum(1 for v in vals if v.strip())
            keep.append(non_empty > 0)
        cleaned = [
            [cell for ci, cell in enumerate(row) if ci < len(keep) and keep[ci]]
            for row in cleaned
        ]

    # 3. 행이 1개 이하 → 표가 아님
    if len(cleaned) <= 1:
        return []
    # 열 제거 후 1열만 남아도 버림
    if max((len(r) for r in cleaned), default=0) < 2:
        return []

    # 4. 반복 헤더 제거 (첫 행 = 헤더 기준)
    header = tuple(cleaned[0])
    deduped = [cleaned[0]]
    for row in cleaned[1:]:
        if tuple(row) != header:
            deduped.append(row)
    return deduped


def _is_valid_table(table_data: list) -> bool:
    """
    추출된 표 데이터가 실제 표인지 엄격하게 검사합니다.

    필터 조건 (하나라도 해당하면 False):
      - 행 수 < 5        : 너무 작은 표 = 오탐 가능성 높음
      - 열 수 < 3        : 1~2열 표 = 텍스트박스 오인식
      - 열 수 > 12       : phantom columns (레이아웃 요소)
      - 불릿 문자 시작 셀 존재 : 본문 단락을 표로 오인식
      - 빈 셀 비율 > 85% : 대부분 비어있는 레이아웃 요소
      - 유니크 값 < 3    : 모두 같은 값 (반복 레이블 등)
      - 50자 초과 셀 ≥ 5 : 단락 텍스트를 표로 오인식
      - 헤더 없음        : 첫 행 모든 셀이 비어있거나 너무 긺
    """
    if not table_data:
        return False

    rows = len(table_data)
    cols = max((len(row) for row in table_data), default=0)

    # 행/열 수 조건
    if rows < 5:
        return False
    if cols < 3:
        return False
    if cols > 12:
        return False

    # 불릿 문자로 시작하는 셀 = 본문 단락 오인식
    BULLET_CHARS = ('▶', '•', '■', '▪', '→', '·', '◆', '○', '●')
    for row in table_data:
        for cell in row:
            cell_str = str(cell or '').strip()
            if any(cell_str.startswith(b) for b in BULLET_CHARS):
                return False

    total_cells = sum(len(row) for row in table_data)
    if total_cells == 0:
        return False
    empty_cells = sum(
        1 for row in table_data for cell in row
        if not str(cell or "").strip()
    )
    if empty_cells / total_cells > 0.85:
        return False

    flat = [str(c).strip() for row in table_data for c in row if str(c or "").strip()]
    if len(set(flat)) < 3:
        return False

    long_cells = sum(
        1 for row in table_data for cell in row
        if cell and len(str(cell)) > 50
    )
    if long_cells > 5:
        return False

    # 첫 행(헤더) 검사: 유효한 헤더 셀이 1개 이상 있어야 함
    # 헤더 셀 = 비어있지 않고 30자 이하 (긴 문장은 헤더가 아님)
    header_row = table_data[0]
    valid_header_cells = sum(
        1 for cell in header_row
        if cell and 0 < len(str(cell).strip()) <= 30
    )
    if valid_header_cells < 2:
        return False

    return True


def _bbox_overlap_ratio(bbox_a: tuple, bbox_b: tuple) -> float:
    """두 bbox의 겹침 비율(IoU)을 반환합니다."""
    x0 = max(bbox_a[0], bbox_b[0])
    y0 = max(bbox_a[1], bbox_b[1])
    x1 = min(bbox_a[2], bbox_b[2])
    y1 = min(bbox_a[3], bbox_b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    inter = (x1 - x0) * (y1 - y0)
    area_a = (bbox_a[2] - bbox_a[0]) * (bbox_a[3] - bbox_a[1])
    area_b = (bbox_b[2] - bbox_b[0]) * (bbox_b[3] - bbox_b[1])
    union = area_a + area_b - inter
    return inter / union if union > 0 else 0.0


def _extract_tables_dual_strategy(plumber_page, debug: bool = False):
    """
    선 기반(lines) 전략만 사용해 표를 추출합니다.
    텍스트 정렬 기반(text) 전략은 불릿 단락을 표로 오인식하므로 비활성화.

    Returns
    -------
    List[Tuple[list, tuple]] — (2D 배열, bbox) 리스트
    """
    settings_a = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}

    try:
        raw_a = plumber_page.find_tables(table_settings=settings_a) or []
    except Exception as exc:
        raw_a = []
        if debug:
            print(f"  [전략A 오류] {exc}")

    if debug:
        print(f"  [전략A-lines] {len(raw_a)}개 감지 (text 전략 비활성화)")

    result = []
    for tbl_a in raw_a:
        extracted = tbl_a.extract()
        if extracted:
            result.append((extracted, tbl_a.bbox))

    return result


def _extract_page_tables(plumber_page, debug: bool = False) -> dict:
    """
    pdfplumber로 한 페이지에서 표를 추출합니다.

    Returns
    -------
    dict:
      tables              : List[List[List[str]]] — 페이지 내 표 목록(행×열)
      table_bboxes        : List[Tuple] — 각 표의 (x0, top, x1, bottom) PDF 좌표
      text_without_tables : str — 표 영역·헤더(상단 20%) 제외 본문 텍스트
    """
    tables: list = []
    table_bboxes: list = []

    # ── 이중 전략 추출 + 유효성 검사 + 후처리 ─────────────
    try:
        for extracted, bbox in _extract_tables_dual_strategy(plumber_page, debug=debug):
            if not _is_valid_table(extracted):
                if debug:
                    rows = len(extracted)
                    cols = max((len(r) for r in extracted), default=0)
                    print(f"  [제외] 유효성 검사 실패 {rows}행×{cols}열 bbox={bbox}")
                continue
            processed = _postprocess_table(extracted)
            if processed:
                tables.append(processed)
                table_bboxes.append(bbox)
    except Exception as exc:
        if debug:
            print(f"  [_extract_page_tables] 오류: {exc}")

    # ── 표·헤더 영역 제외 텍스트 추출 ──────────────────────
    text_without_tables = ""
    try:
        header_thresh_y = plumber_page.height * 0.20
        words = plumber_page.extract_words() or []
        non_table_words = []
        for word in words:
            # 헤더 영역(상단 20%) 제외
            if word["top"] < header_thresh_y:
                continue
            # 표 영역 제외 (3px 여유)
            in_table = any(
                bbox[0] - 3 <= word["x0"] and word["x1"] <= bbox[2] + 3
                and bbox[1] - 3 <= word["top"] and word["bottom"] <= bbox[3] + 3
                for bbox in table_bboxes
            )
            if not in_table:
                non_table_words.append(word["text"])
        text_without_tables = " ".join(non_table_words)
    except Exception as exc:
        if debug:
            print(f"  [_extract_page_tables] 텍스트 추출 오류: {exc}")

    if debug and tables:
        for idx, tbl in enumerate(tables):
            cols = max((len(r) for r in tbl), default=0)
            print(f"  [표 {idx+1}] {len(tbl)}행 × {cols}열  bbox={table_bboxes[idx]}")

    return {
        "tables":              tables,
        "table_bboxes":        table_bboxes,
        "text_without_tables": text_without_tables,
    }
