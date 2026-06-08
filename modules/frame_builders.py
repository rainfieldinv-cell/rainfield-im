"""
frame_builders.py
─────────────────────────────────────────────────────────
"틀 먼저 → 채움" 방식의 본문 슬라이드 빌더 + 페이지 레이아웃 자동 판별기.

build_full_presentation(_build_content_block)이 변환 중 페이지마다
classify_page()로 유형을 자동 판별하고, 우리가 구현한 유형(E)은
이 모듈의 전용 빌더로, 아직 미구현 유형은 기존 build_content_slide로
폴백(fallback)합니다.

유형 (rainfield-layout-patterns 참고):
  A. 전체폭 데이터 표              (미구현 → 폴백)
  B. 좌우 [이미지 표 | 텍스트 표]  (미구현 → 폴백)
  C. 좌우 [이미지 | 비교 표]       (미구현 → 폴백)
  D. 표/차트                       (미구현 → 폴백)
  E. 데이터 표 + 이미지 표         (구현 — build_frame_e_slide)

좌표는 사람이 만든 제안서(넷마블 G타워) 슬9 실측값에 맞춤:
  좌 데이터표 L0.43 W4.99 / 우 이미지표 L5.56 W4.85 / 본문 T1.83
"""

import io
from collections import Counter

from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from modules import page_builders
from modules.page_builders import (
    clone_slide_layout,
    _replace_footer_business_name,
    _replace_text_frame_content,
    auto_resize_text_to_fit,
)
from modules.ai_slide_builders import (
    style_table, PALETTE, _replace_text_keep_runs,
)

# 이미지 행 라벨(데이터표에서 제외하고 이미지표 라벨로 사용)
IMG_LABELS = {"위치도", "조감도", "조 감 도", "위 치 도", "투시도", "전경", "건물전경", "건물 전경"}

# 유형 E 판별 임계값
_E_MIN_DATA_ROWS = 5       # label-value 표 최소 행 수
_E_MIN_IMG_AREA  = 8000    # 사진으로 인정할 최소 픽셀 면적(px*px)
_E_MAX_IMGS      = 3       # 이미지 표에 넣을 최대 사진 수


# ──────────────────────────────────────────────────────
# 표 정규화 / 판별 유틸
# ──────────────────────────────────────────────────────
def _to_label_value(rows):
    """2D 표 → [[라벨, 값], ...]. 행마다 비어있지 않은 첫 셀=라벨, 나머지 join=값."""
    out = []
    for row in rows or []:
        cells = [str(c or "").strip() for c in row]
        ne = [c for c in cells if c and c != "•"]
        if not ne:
            continue
        out.append([ne[0], " ".join(ne[1:]).replace("•", "").strip()])
    return out


def _is_label_value(rows):
    """행 대부분(60%+)이 비어있지않은 셀 2개 이하 → 구분|내용 형태로 간주."""
    if not rows:
        return False
    le2 = 0
    for r in rows:
        ne = [c for c in r if str(c or "").strip() and str(c or "").strip() != "•"]
        if len(ne) <= 2:
            le2 += 1
    return le2 >= max(2, len(rows) * 0.6)


def _big_images(images):
    """로고·아이콘 제외한 사진급 이미지만."""
    return [im for im in (images or [])
            if (im.get("width", 0) * im.get("height", 0)) >= _E_MIN_IMG_AREA]


def _main_table(tables):
    return max(tables, key=lambda t: len(t) * max((len(r) for r in t), default=0),
               default=None)


def classify_page(page: dict) -> str:
    """
    파싱된 페이지 dict(tables/images 포함) → 레이아웃 유형 문자열.

    현재 'E'만 전용 빌더가 있고 나머지는 'fallback'(기존 빌더)으로 보냅니다.
    """
    tables = page.get("tables") or []
    images = page.get("images") or []
    if not tables:
        return "fallback"
    main = _main_table(tables)
    if not main or len(main) < _E_MIN_DATA_ROWS:
        return "fallback"
    big = _big_images(images)
    # 유형 E: 큰 사진 1~3장 + 구분|내용(label-value) 데이터 표
    if 1 <= len(big) <= _E_MAX_IMGS and _is_label_value(main):
        return "E"
    return "fallback"


# ──────────────────────────────────────────────────────
# 헤더(제목/부제목/본문) 교체 — build_content_slide와 동일 규칙
# ──────────────────────────────────────────────────────
def _fill_header(slide, title, subtitle, intro):
    LEFT_MARGIN_CM = 1.09
    TOP_LIMIT_CM   = 4.0
    TOLERANCE      = Cm(0.5)
    header_shapes = sorted(
        [sh for sh in slide.shapes
         if sh.has_text_frame
         and abs(sh.left - Cm(LEFT_MARGIN_CM)) < TOLERANCE
         and sh.top / 360000 < TOP_LIMIT_CM],
        key=lambda s: s.top,
    )
    if len(header_shapes) >= 1:
        _replace_text_frame_content(header_shapes[0].text_frame, title)
    if len(header_shapes) >= 2:
        _replace_text_frame_content(header_shapes[1].text_frame, subtitle)
    if len(header_shapes) >= 3:
        if intro and intro.strip():
            _replace_text_frame_content(header_shapes[2].text_frame, intro.strip())
            auto_resize_text_to_fit(header_shapes[2].text_frame, max_size=10.5, min_size=8.0)
        else:
            _replace_text_frame_content(header_shapes[2].text_frame, "")


# ──────────────────────────────────────────────────────
# 유형 E 빌더 — [데이터 표 | 이미지 표]
# ──────────────────────────────────────────────────────
def build_frame_e_slide(prs, *, title: str, subtitle: str, intro: str,
                        data_rows: list, imgs: list, img_labels: list = None,
                        business_name: str = "", sub_label: str = None):
    """
    유형 E 슬라이드: 왼쪽 구분|내용 데이터 표, 오른쪽 구분|내용 이미지 표.

    Parameters
    ----------
    data_rows  : [[라벨, 값], ...]  (행 0 = 헤더 "구분"/"내용")
    imgs       : [{"data"|"bytes": bytes, "width"|"w": int, "height"|"h": int}, ...]
    img_labels : 이미지 표 각 행 라벨 (없으면 ["조감도","위치도",...])
    sub_label  : 표 위 소제목바 텍스트 (None이면 생략)
    """
    slide = clone_slide_layout(prs, "content", skip_graphic_frames=True)
    _fill_header(slide, title, subtitle, intro)

    # ── 소제목바 (선택) ──
    if sub_label:
        sub = slide.shapes.add_textbox(Inches(0.43), Inches(1.5), Inches(3.0), Inches(0.3))
        run = sub.text_frame.paragraphs[0].add_run()
        run.text = sub_label
        run.font.name = "피플폰트 Bold"
        run.font.size = Pt(12)
        run.font.color.rgb = PALETTE["navy_dark"]

    # ── 왼쪽: 데이터 표 (L0.43 W4.99) ──
    rows = max(1, len(data_rows))
    gL = slide.shapes.add_table(rows, 2, Inches(0.43), Inches(1.83),
                                Inches(4.99), Inches(rows * 0.45))
    tL = gL.table
    tL.columns[0].width = Inches(1.40)
    tL.columns[1].width = Inches(3.59)
    for ri, pair in enumerate(data_rows):
        lab = pair[0] if len(pair) > 0 else ""
        val = pair[1] if len(pair) > 1 else ""
        _replace_text_keep_runs(tL.cell(ri, 0).text_frame, lab)
        _replace_text_keep_runs(tL.cell(ri, 1).text_frame, val)
        if ri >= 1:
            from pptx.enum.text import PP_ALIGN
            tL.cell(ri, 1).text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    style_table(gL, has_header=True, label_cols=(0,),
                header_fill=PALETTE["navy_dark"], label_fill=PALETTE["label_gray"])

    # ── 오른쪽: 이미지 표 (L5.56 W4.85) ──
    labels = img_labels or ["조감도", "위치도", "투시도"]
    img_rows = max(1, len(imgs))
    R_L, R_T, R_W = 5.56, 1.83, 4.85
    ROW_H = 2.30
    gR = slide.shapes.add_table(img_rows + 1, 2, Inches(R_L), Inches(R_T),
                                Inches(R_W), Inches(0.4 + img_rows * ROW_H))
    tR = gR.table
    tR.columns[0].width = Inches(0.95)
    tR.columns[1].width = Inches(3.90)
    _replace_text_keep_runs(tR.cell(0, 0).text_frame, "구 분")
    _replace_text_keep_runs(tR.cell(0, 1).text_frame, "내 용")
    tR.rows[0].height = Inches(0.4)
    for i in range(img_rows):
        _replace_text_keep_runs(tR.cell(i + 1, 0).text_frame,
                                labels[i] if i < len(labels) else "사진")
        tR.rows[i + 1].height = Inches(ROW_H)
    style_table(gR, has_header=True, label_cols=(0,),
                header_fill=PALETTE["navy_dark"], label_fill=PALETTE["label_gray"])

    # ── 이미지 삽입 (내용칸 위 비율유지·중앙) ──
    content_L = R_L + 0.95
    content_W = 3.90
    pad = 0.1
    for i, im in enumerate(imgs):
        data = im.get("data") or im.get("bytes")
        iw0  = im.get("width") or im.get("w") or 1
        ih0  = im.get("height") or im.get("h") or 1
        if not data:
            continue
        cell_T = R_T + 0.4 + i * ROW_H
        boxL, boxT = content_L + pad, cell_T + pad
        boxW, boxH = content_W - 2 * pad, ROW_H - 2 * pad
        scale = min(boxW / iw0, boxH / ih0)
        iw, ih = iw0 * scale, ih0 * scale
        px = boxL + (boxW - iw) / 2
        py = boxT + (boxH - ih) / 2
        slide.shapes.add_picture(io.BytesIO(data), Inches(px), Inches(py),
                                 Inches(iw), Inches(ih))

    _replace_footer_business_name(slide, business_name)
    return slide


# ──────────────────────────────────────────────────────
# 라우터 진입점 — _build_content_block에서 호출
# ──────────────────────────────────────────────────────
def build_page_auto(prs, page: dict, *, title: str, subtitle: str,
                    body_text: str, business_name: str) -> bool:
    """
    페이지 유형을 자동 판별해 틀-우선 빌더로 슬라이드를 만든다.
    LLM 토글(env RAINFIELD_LLM=1)이 켜져 있고 원문이 있으면 LLM 구조화 경로 우선.
    아니면 유형 E면 전용 E빌더, 그 외 모든 페이지는 범용 골격 빌더로 처리.
    처리했으면 True(항상), 예외 시에만 호출측이 기존 빌더로 폴백.
    """
    raw = page.get("raw_text", "") or ""
    sec = page.get("_sec_int")

    # ── 고정 페이지: 투자구조도(도형으로 직접 그림) ──
    if page.get("_invest_diagram") is not None:
        build_invest_diagram_slide(
            prs, page.get("_invest_diagram") or {},
            section_label=page.get("section_label"),
            subtitle=page.get("subtitle"),
            intro=page.get("_invest_intro", ""),
            business_name=business_name)
        return True

    # ── 고정 페이지: 투자구조도(원본 구조도 영역 이미지 캡처 — 비활성) ──
    if page.get("_invest_image"):
        build_invest_image_slide(
            prs,
            section_label=page.get("section_label"),
            subtitle=page.get("subtitle"),
            intro=page.get("_invest_intro", ""),
            img=page.get("_invest_image"),
            business_name=business_name)
        return True

    # ── 고정 페이지: [섹션1] 사모사채 개요 → 전용 고정 빌더(구분열 불변, 내용만 LLM) ──
    if sec == 1 and raw.strip():
        try:
            from modules.ai_slide_builders import (
                generate_sasae_overview, build_slide_5_sasae_overview)
            res = generate_sasae_overview(raw)
            if res.get("ok") and res.get("data"):
                sl = build_slide_5_sasae_overview(prs, res["data"], business_name=business_name)
                # ★사모사채개요 표 글꼴 전부 Bold 통일(뒤죽박죽 방지 — 사용자 지시)
                try:
                    for sh in sl.shapes:
                        if sh.has_table:
                            for row in sh.table.rows:
                                for cell in row.cells:
                                    for pp in cell.text_frame.paragraphs:
                                        _set_para_default_font(pp, 10.5, _FONT_BOLD)
                                        for rn in pp.runs:
                                            rn.font.name = _FONT_BOLD
                                            rn.font.size = Pt(10.5)
                except Exception:
                    pass
                return True
        except Exception as _exc:
            print(f"[build_page_auto] 사모사채개요 고정빌더 실패 → 구조화 폴백: {_exc}")

    # enrich_and_number()로 미리 구조화된 결과가 있으면 사용
    struct = page.get("_struct")
    if isinstance(struct, dict):
        build_structured_slide(
            prs, struct, business_name=business_name,
            section_label=page.get("section_label"),
            subtitle=page.get("subtitle"),
            images=page.get("images"),
            red_texts=page.get("_red_texts"),
            underline_texts=page.get("_underline_texts"))
        return True

    ptype = classify_page(page)

    if ptype == "E":
        tables = page.get("tables") or []
        main = _main_table(tables)
        norm = _to_label_value(main)

        def _norm_key(s):
            return s.replace(" ", "")
        img_label_set = {_norm_key(x) for x in IMG_LABELS}
        img_labels_found = [r[0] for r in norm if _norm_key(r[0]) in img_label_set]
        data_rows = [r for r in norm if _norm_key(r[0]) not in img_label_set]
        imgs = _big_images(page.get("images") or [])[: _E_MAX_IMGS]
        if data_rows and imgs:
            build_frame_e_slide(
                prs, title=title, subtitle=subtitle, intro=body_text,
                data_rows=data_rows, imgs=imgs,
                img_labels=img_labels_found or None,
                business_name=business_name,
            )
            return True

    # 그 외 모든 페이지 → 범용 골격 빌더
    build_universal_content_slide(
        prs, title=title, subtitle=subtitle, intro=body_text,
        tables=page.get("tables") or [], business_name=business_name,
    )
    return True


# ──────────────────────────────────────────────────────
# 범용 골격 빌더 — 대전 제안서 실측 골격 재현
#   섹션라벨/소제목(템플릿 자리) + 인트로(별도 글상자) + 표(정규화·스택) + 푸터/출처
#   ★본문은 절대 제목/소제목 자리에 넣지 않고 새 글상자로. 글상자 내부여백 0.
# ──────────────────────────────────────────────────────
_BLACK = PALETTE["black"]
_FONT_LIGHT = "피플폰트 Light"
_FONT_BOLD  = "피플폰트 Bold"

# 대전 실측 좌표(in)
_INTRO_L, _INTRO_T, _INTRO_W = 0.43, 1.01, 8.65
_BODY_TOP_START = 1.55      # 표/본문 영역 시작
_BODY_BOTTOM    = 6.85      # 각주/푸터 전 한계
_TBL_L, _TBL_W  = 0.43, 9.97
_SRC_L, _SRC_T, _SRC_W = 0.43, 7.19, 8.65


def _est_text_height(text, W, size):
    """글상자 높이(in) 추정 — 자동줄바꿈 고려. 빈 줄 군더더기 없이 내용에 딱 맞춤."""
    char_w = size * 0.78 / 72.0           # 한글 기준 글자폭 근사
    cpl = max(1, int(W / char_w))
    line_h = size * 1.32 / 72.0           # 한 줄 높이(여유 최소화)
    lines = 0
    for ln in str(text).split("\n"):
        lines += max(1, -(-len(ln) // cpl))   # ceil
    return max(line_h, lines * line_h) + 0.03


def _emph_segments(line, red_set, ul_set=None):
    """한 줄을 [(텍스트, is_red, is_underline)] 조각으로 분할.
       원본의 빨간 글씨/밑줄 부분만 강조(겹쳐도 문자 단위 마스크로 정확히 처리)."""
    n = len(line)
    if n == 0:
        return [("", False, False)]
    red_mask = [False] * n
    ul_mask = [False] * n

    def _mark(mask, frags, minlen):
        for f in sorted({str(s).strip() for s in (frags or []) if len(str(s).strip()) >= minlen},
                        key=len, reverse=True):
            start = 0
            while True:
                idx = line.find(f, start)
                if idx == -1:
                    break
                for k in range(idx, idx + len(f)):
                    mask[k] = True
                start = idx + len(f)

    _mark(red_mask, red_set, 4)   # 빨강은 4글자 이상(루시드 등 과다 방지)
    _mark(ul_mask, ul_set, 2)     # 밑줄은 2글자 이상
    if not any(red_mask) and not any(ul_mask):
        return [(line, False, False)]
    segs, i = [], 0
    while i < n:
        r, u = red_mask[i], ul_mask[i]
        j = i
        while j < n and red_mask[j] == r and ul_mask[j] == u:
            j += 1
        segs.append((line[i:j], r, u))
        i = j
    return segs


def _add_textbox(slide, L, T, W, H, text, *, size=10.5, bold=False,
                 color=None, align=PP_ALIGN.LEFT, red_set=None, ul_set=None):
    """글상자 생성 — 내부여백 0, 텍스트 내용에 딱 맞는 높이(군더더기 여백 없음). 반환=(shape, 높이in).
       red_set/ul_set 지정 시 원본 빨간 글씨/밑줄에 해당하는 부분만 빨강·밑줄로 표시."""
    text = str(text).rstrip("\n ")        # 끝의 빈 줄/공백 제거(빈 여백 방지)
    est_h = _est_text_height(text, W, size)
    tb = slide.shapes.add_textbox(Inches(L), Inches(T), Inches(W), Inches(est_h))
    tf = tb.text_frame
    tf.word_wrap = True
    # 자동맞춤 끔 — PPT가 안 줄여주는 빈 여백 문제 방지(계산한 높이 그대로 사용)
    try:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    except Exception:
        pass
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    nm = _FONT_BOLD if bold else _FONT_LIGHT
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.font.size = Pt(size)                       # 문단 기본(빈 줄에도 적용)
        p.font.name = nm
        for seg, is_red, is_ul in _emph_segments(line, red_set, ul_set):
            run = p.add_run()
            run.text = seg
            run.font.name = nm
            run.font.size = Pt(size)
            run.font.color.rgb = _RED if is_red else (color or _BLACK)
            if is_ul:
                run.font.underline = True
    return tb, est_h


def _drop_empty_cols(rows):
    """완전히 빈 열(phantom column) 제거."""
    if not rows:
        return rows
    ncol = max(len(r) for r in rows)
    keep = [c for c in range(ncol)
            if any(str((r[c] if c < len(r) else "") or "").strip() for r in rows)]
    return [[(r[c] if c < len(r) else "") for c in keep] for r in rows]


def _looks_like_prose(rows):
    """표가 아니라 문장이 칸칸이 쪼개진 가짜 표인지 판정."""
    norm = _to_label_value(rows)
    if not norm:
        return True
    empty_first = sum(1 for r in rows if not str((r[0] if r else "") or "").strip())
    long_labels = sum(1 for lab, _ in norm if len(lab) > 18)
    return empty_first > len(rows) * 0.4 or long_labels > len(norm) * 0.5


def _add_norm_table(slide, rows, L, T, W):
    """정규화 표 삽입 → (graphicFrame, 사용높이in). 칸 너비를 내용 형태에 맞춤."""
    label_value = _is_label_value(rows)
    if label_value:
        data = _to_label_value(rows)
        ncol = 2
    else:
        data = _drop_empty_cols(rows)
        ncol = max((len(r) for r in data), default=0)
    if not data or ncol == 0:
        return None, 0.0

    nrow = len(data)
    row_h = 0.34
    height = max(row_h, nrow * row_h)
    gf = slide.shapes.add_table(nrow, ncol, Inches(L), Inches(T),
                                Inches(W), Inches(height))
    t = gf.table

    if label_value:
        lab_w = min(1.6, W * 0.24)
        t.columns[0].width = Inches(lab_w)
        t.columns[1].width = Inches(W - lab_w)
        for ri, pair in enumerate(data):
            _replace_text_keep_runs(t.cell(ri, 0).text_frame, pair[0])
            _replace_text_keep_runs(t.cell(ri, 1).text_frame,
                                    pair[1] if len(pair) > 1 else "")
            t.cell(ri, 1).text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        style_table(gf, has_header=False, label_cols=(0,),
                    label_fill=PALETTE["label_gray"])
    else:
        widths = _content_col_widths(data, ncol, W, 10.5)
        for c in range(ncol):
            t.columns[c].width = Inches(widths[c])
        for ri, row in enumerate(data):
            for ci in range(ncol):
                val = str((row[ci] if ci < len(row) else "") or "").strip()
                _replace_text_keep_runs(t.cell(ri, ci).text_frame, val)
        style_table(gf, has_header=True, label_cols=(),
                    header_fill=PALETTE["navy_dark"])
    return gf, height


def build_universal_content_slide(prs, *, title: str, subtitle: str, intro: str,
                                  tables: list, business_name: str = "",
                                  source: str = None):
    """대전 실측 골격을 재현하는 범용 본문 슬라이드."""
    slide = clone_slide_layout(prs, "content", skip_graphic_frames=True)

    # 1) 섹션라벨/소제목은 템플릿 자리에, 본문 자리(plh[2])는 비움
    _fill_header(slide, title, subtitle, "")

    # 2) 인트로 = 별도 글상자 (Light 10.5, 여백0). ★제목칸에 안 넣음
    top = _BODY_TOP_START
    if intro and intro.strip():
        _add_textbox(slide, _INTRO_L, _INTRO_T, _INTRO_W, 0.45,
                     intro.strip(), size=10.5, bold=False)
        top = max(top, _INTRO_T + 0.55)

    # 3) 표 — 가짜표(문장쪼개짐) 제외, 정규화해 세로 스택
    real_tables = [tb for tb in (tables or []) if tb and not _looks_like_prose(tb)]
    for tb in real_tables:
        if top > _BODY_BOTTOM - 0.5:
            break   # TODO: 페이지 분할 (1/2)(2/2) — 후속 작업
        avail = _BODY_BOTTOM - top
        gf, used = _add_norm_table(slide, tb, _TBL_L, top, _TBL_W)
        if gf is None:
            continue
        # 표가 남은 공간보다 크면 그대로 두되 다음 표는 폴백 한계에 막힘
        top += min(used, avail) + 0.25

    # 4) 푸터(사업명) + 출처 글상자(여백0)
    _replace_footer_business_name(slide, business_name)
    if source and source.strip():
        _add_textbox(slide, _SRC_L, _SRC_T, _SRC_W, 0.22,
                     f"* 출처 : {source.strip()}", size=9,
                     color=PALETTE["gray_text"])
    return slide


# ──────────────────────────────────────────────────────
# 구조화(LLM) 슬라이드 빌더 — llm_structure.structure_page() 결과를 렌더
# ──────────────────────────────────────────────────────
_TOTAL_KW = ("합계", "소계", "총계", "합 계", "총 계", "총 합계", "총합계")


def _classify_total_rows(data, ncol):
    """행 인덱스를 (소계행, 합계/총합계행) 두 집합으로 분류.
       ★'라벨 셀'(짧은 칸 ≤12자)에 키워드가 있을 때만 인정 — 긴 값 속 '합계'(예: 연면적
         '지상…/지하…/합계 172,…')는 제외. 소계는 '소계'만, 합계/총계/총합계는 grand-total."""
    subtotal, grandtotal = set(), set()
    for i, row in enumerate(data):
        cells = [str((row[c] if c < len(row) else "") or "").strip() for c in range(ncol)]
        is_sub = any(len(cx) <= 12 and any(k in cx for k in ("소계", "소 계")) for cx in cells)
        #  '총…'(총매출/총비용/총계/총합계 등 짧은 합산 라벨)도 합계행으로 인정
        is_grand = any(len(cx) <= 12 and (cx.startswith("총")
                       or any(k in cx for k in ("합계", "총계", "총합계", "합 계", "총 계")))
                       for cx in cells)
        if is_sub:
            subtotal.add(i)
        elif is_grand:
            grandtotal.add(i)
    return subtotal, grandtotal


def _merge_vertical_runs(table, data, ncol, header_rows=0):
    """열에서 값 있는 칸 아래로 이어지는 빈 칸들을 세로 병합(반복값=첫 칸만, 아래 빈칸 → 병합).
       ★구분열(ci==0)은 소계 행을 넘어 카테고리를 이어 병합(원본: '공동주택'이 소계까지 한 칸).
       ★합계/총합계 행은 모든 열에서 병합 경계 — 위 값이 합계행으로 번지지 않음."""
    nrow = len(data)
    subtotal_rows, grandtotal_rows = _classify_total_rows(data, ncol)
    for ci in range(ncol):
        # 구분열만 소계를 관통해 병합, 나머지 열은 소계·합계 모두 경계
        boundary = grandtotal_rows if ci == 0 else (grandtotal_rows | subtotal_rows)
        ri = header_rows
        while ri < nrow:
            if ri in boundary:          # 경계행은 병합 시작도 안 함
                ri += 1
                continue
            val = str((data[ri][ci] if ci < len(data[ri]) else "") or "").strip()
            if not val:
                ri += 1
                continue
            end = ri
            for rj in range(ri + 1, nrow):
                if rj in boundary:      # 경계행 만나면 중단
                    break
                v2 = str((data[rj][ci] if ci < len(data[rj]) else "") or "").strip()
                if v2 == "":
                    end = rj
                else:
                    break
            if end > ri:
                try:
                    table.cell(ri, ci).merge(table.cell(end, ci))
                except Exception:
                    pass
            ri = end + 1


def _merge_total_rows(table, data, ncol):
    """합계/소계/총매출 등 합산행에서 라벨을 빈 칸들에 가로 병합
       (예 '터미널 합계(A)'·'총 매출 (A)'가 옆 빈 칸들을 차지)."""
    subtotal, grandtotal = _classify_total_rows(data, ncol)
    total_set = subtotal | grandtotal
    for ri, row in enumerate(data):
        if ri not in total_set:
            continue
        txts = [str((row[c] if c < len(row) else "") or "").strip() for c in range(ncol)]
        first = next((i for i, tx in enumerate(txts) if tx), None)
        if first is None:
            continue
        end = first
        for j in range(first + 1, ncol):
            if txts[j] == "":
                end = j
            else:
                break
        if end > first:
            try:
                table.cell(ri, first).merge(table.cell(ri, end))
            except Exception:
                pass


def _font_for_rowh(row_h):
    if row_h >= 0.30:
        return 10.5
    if row_h >= 0.24:
        return 9
    if row_h >= 0.19:
        return 8
    if row_h >= 0.155:
        return 7
    return 6


def _uniform_font_rowh(nrow, avail):
    """표 전체에 통일 적용할 (글자크기, 행높이). 10.5pt 우선, 안 들어가면 통일로 낮춤(최저 8pt)."""
    for fp in (10.5, 10, 9, 8):
        rh = fp * 1.5 / 72.0 + 0.05   # 한 줄 행높이 근사(in)
        if nrow * rh <= avail:
            return fp, rh
    return 8, 8 * 1.5 / 72.0 + 0.05   # 최저 8pt(넘치면 페이지분할 대상)


def _set_para_default_font(p, font_pt, name):
    """문단 기본 글꼴(defRPr) 크기·글꼴(라틴+한글 ea) 지정 — ★빈 칸에도 적용되어
       빈 칸이 큰 기본 글씨로 행높이를 키우는 문제 제거."""
    p.font.size = Pt(font_pt)
    p.font.name = name                      # a:latin
    sz = str(int(font_pt * 100))
    try:
        defRPr = p._p.get_or_add_pPr().get_or_add_defRPr()
        defRPr.set('sz', sz)
        ea = defRPr.find(qn('a:ea'))
        if ea is None:
            ea = defRPr.makeelement(qn('a:ea'), {})
            defRPr.append(ea)
        ea.set('typeface', name)
        # 빈 문단의 줄 높이를 좌우하는 endParaRPr 크기도 축소
        endpr = p._p.find(qn('a:endParaRPr'))
        if endpr is None:
            endpr = p._p.makeelement(qn('a:endParaRPr'), {})
            p._p.append(endpr)
        endpr.set('sz', sz)
    except Exception:
        pass


def _compact_grid(table, font_pt, row_h, kind="grid", has_header=False):
    """표 글자크기·행높이 통일 + Bold/Light.
       - grid: 줄바꿈 OFF(한 줄)·통일 행높이.
       - label_value: 내용 칸(ci=1)은 줄바꿈 ON + 행높이 가변(row_h가 리스트면 행별 높이).
       ★헤더행·구분열(label_value col0)은 Bold, 나머지 Light — 빈 칸 포함 모든 칸에 적용."""
    heights = row_h if isinstance(row_h, (list, tuple)) else None
    for ri, row in enumerate(table.rows):
        if heights is not None:
            row.height = Inches(heights[ri] if ri < len(heights) else heights[-1])
        elif has_header and ri == 0:
            row.height = Cm(0.6)        # ★표 헤더행은 0.6cm, 내용행은 축소(사진 공간 확보)
        else:
            row.height = Inches(row_h)
        for ci, cell in enumerate(row.cells):
            is_header_cell = has_header and ri == 0
            is_bold = is_header_cell or (kind == "label_value" and ci == 0)
            name = _FONT_BOLD if is_bold else _FONT_LIGHT
            cell.margin_left = cell.margin_right = Inches(0.04)
            cell.margin_top = cell.margin_bottom = Inches(0.02 if kind == "label_value" else 0)
            tf = cell.text_frame
            # label_value 내용칸(긴 약정문)만 줄바꿈 허용. 세로정렬은 전부 가운데 통일(사용자 지시)
            if kind == "label_value" and ci == 1 and not is_header_cell:
                tf.word_wrap = True
            else:
                tf.word_wrap = False
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in tf.paragraphs:
                # ★grid 셀 정렬 일관 통일(가운데) — 템플릿 빈칸 상속으로 인한 뒤죽박죽 정렬 방지
                if kind == "grid":
                    p.alignment = PP_ALIGN.CENTER
                _set_para_default_font(p, font_pt, name)   # 빈 칸 포함
                for run in p.runs:                          # 채워진 칸: 크기·글꼴 통일(헤더 Bold 보장)
                    run.font.size = Pt(font_pt)
                    run.font.name = name


_RED = RGBColor(0xC0, 0x00, 0x00)


def _cell_is_red(txt, red_set):
    """셀 텍스트가 원본 빨간 글씨와 매칭되면 True.
       ★셀 텍스트가 '빨간 주석 안에 포함'된 경우(예: 검정 '루시드' ⊂ 빨간 주석)는 제외.
         셀 == 빨간조각, 또는 빨간조각이 셀의 큰 부분일 때만 빨강 처리."""
    t = (txt or "").strip()
    if len(t) < 2 or not red_set:
        return False
    for r in red_set:
        r = (r or "").strip()
        if len(r) < 2:
            continue
        if r == t:
            return True
        # 빨간 조각이 셀 안에 있고, 셀 길이의 절반 이상을 차지하면 강조로 인정
        if r in t and len(r) >= max(2, len(t) * 0.5):
            return True
    return False


def _set_cell_red(cell):
    """셀 글자만 빨강 (원본은 '빨간 글씨' — 테두리 금지). 빨간 테두리는 _set_cell_red_border 별도."""
    for p in cell.text_frame.paragraphs:
        for run in p.runs:
            run.font.color.rgb = _RED


def _set_cell_red_border(cell):
    """셀 4변 빨간 테두리 1pt (원본 셀 외곽이 빨간 경우 — 시세표 '본건' 행 등)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
        ln = tcPr.makeelement(qn(tag), {"w": "12700", "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        sf = ln.makeelement(qn("a:solidFill"), {})
        sf.append(sf.makeelement(qn("a:srgbClr"), {"val": "C00000"}))
        ln.append(sf)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
        tcPr.insert(i, ln)


def _grid_font(ncol):
    """표 글자크기 — 사용자 지시: 모든 표 10.5pt 통일."""
    return 10.5


def _rowh(font_pt):
    return font_pt * 1.32 / 72.0 + 0.02   # 한 줄 내용행 높이(in) — 최대한 축소(사진 공간 확보)


def _set_one_cell_fill_alpha(cell, hex_color, alpha_pct):
    """셀 1개 채움을 hex_color + 투명도(alpha_pct% 불투명)로."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:pattFill"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    sf = tcPr.makeelement(qn("a:solidFill"), {})
    clr = sf.makeelement(qn("a:srgbClr"), {"val": hex_color})
    clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(alpha_pct * 1000))}))
    sf.append(clr)
    ins = 0
    for i, ch in enumerate(tcPr):
        if ch.tag.endswith(("lnL", "lnR", "lnT", "lnB")):
            ins = i + 1
    tcPr.insert(ins, sf)


def _content_col_widths(data, ncol, W, font_pt):
    """★열 너비를 내용 길이에 비례 배분(균등분할 금지). 짧은 칸은 좁게, 긴 칸은 넓게.
       전체 폭 W에 맞춰 스케일(짧은 열의 상대적 좁음은 유지)."""
    char_w = font_pt * 0.78 / 72.0
    maxlen = [1] * ncol
    for row in data:
        for c in range(ncol):
            txt = str(row[c] if c < len(row) else "")
            ll = max((len(s) for s in txt.split("\n")), default=0)
            if ll > maxlen[c]:
                maxlen[c] = ll
    raw = [max(0.45, maxlen[c] * char_w + 0.14) for c in range(ncol)]
    tot = sum(raw)
    scale = W / tot if tot > 0 else 1.0
    return [w * scale for w in raw]


_CORP_KEYS = ("회사명", "업체명", "대표자", "사업자번호", "법인등록번호",
              "설립일", "설립일자", "결산월", "결산일", "법인구분", "업종명", "주소", "본점소재지")


def _is_corp_info(tdef):
    """기업개요(기본정보) label_value 표인지 — 4열(구분|내용|구분|내용)로 바꿀 대상."""
    if (tdef.get("kind") or "") != "label_value":
        return False
    labels = [str((r[0] if r else "") or "") for r in (tdef.get("rows") or [])]
    hits = sum(1 for L in labels if any(k in L for k in _CORP_KEYS))
    return hits >= 3


def _corp_to_4col(tdef):
    """기업개요 label_value(세로 나열) → 4열 grid(구분|내용|구분|내용). 연속 2행을 좌/우로."""
    rows = [[str((r[0] if len(r) > 0 else "") or ""), str((r[1] if len(r) > 1 else "") or "")]
            for r in (tdef.get("rows") or [])]
    rows = [r for r in rows if r[0] or r[1]]
    out = []
    for i in range(0, len(rows), 2):
        left = rows[i]
        right = rows[i + 1] if i + 1 < len(rows) else ["", ""]
        out.append([left[0], left[1], right[0], right[1]])
    return {"title": tdef.get("title", ""), "kind": "grid",
            "header": ["구분", "내용", "구분", "내용"], "rows": out}


def _parse_tdef(tdef):
    """LLM 표 정의 → (kind, header, body_rows, ncol)."""
    kind = (tdef.get("kind") or "").strip()
    if kind == "label_value":
        body = [[str((r[0] if len(r) > 0 else "") or ""),
                 str((r[1] if len(r) > 1 else "") or "")] for r in (tdef.get("rows") or [])]
        body = [r for r in body if r[0] or r[1]]
        # ★옆으로 된 표는 맨 위에 네이비 「구분 | 내용」 헤더행 추가(분할 시 매 장 반복)
        return "label_value", ["구분", "내용"], body, 2
    header = [str(h or "") for h in (tdef.get("header") or [])]
    body = [[str(c or "") for c in r] for r in (tdef.get("rows") or [])]
    ncol = max([len(header)] + [len(r) for r in body]) if (header or body) else 0
    return "grid", header, body, ncol


def _set_header_fill_alpha(table, hex_color, alpha_pct):
    """헤더행(row0) 셀 채움을 hex_color + 투명도(alpha_pct% 불투명)로. 글자=검정(밝은 배경)."""
    M = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for cell in table.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:pattFill"):
            for el in tcPr.findall(qn(tag)):
                tcPr.remove(el)
        sf = tcPr.makeelement(qn("a:solidFill"), {})
        clr = sf.makeelement(qn("a:srgbClr"), {"val": hex_color})
        clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(alpha_pct * 1000))}))
        sf.append(clr)
        # 테두리(lnL/R/T/B) 뒤, fill은 그 다음에 와야 하므로 적절히 삽입
        ins = 0
        for i, ch in enumerate(tcPr):
            if ch.tag.endswith(("lnL", "lnR", "lnT", "lnB")):
                ins = i + 1
        tcPr.insert(ins, sf)
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0, 0, 0)


def _render_table_chunk(slide, kind, header, rows, ncol, L, T, W, font_pt, row_h, red_counts=None,
                        hdr_fill_hex=None, hdr_alpha=None, anchor_rows=None, is_last_chunk=True):
    """헤더(있으면)+행들을 렌더 → 높이(in).
       row_h: 스칼라(그리드 통일 행높이) 또는 본문행별 높이 리스트(label_value 가변).
       red_set: 원본 빨간 글씨 집합 — 매칭 셀은 빨간 글씨+빨간 테두리.
       hdr_fill_hex/hdr_alpha: 헤더행 특수 색(중첩표용 3E95BE+투명도)."""
    data = ([header] if header else []) + rows
    if not data or ncol == 0:
        return 0.0
    nrow = len(data)
    if isinstance(row_h, (list, tuple)):
        hdr_h = _rowh(font_pt)
        row_heights = ([hdr_h] if header else []) + list(row_h)
        height = sum(row_heights)
        row_h = row_heights        # 아래 _compact_grid에 행별 높이 전달
    else:
        height = nrow * row_h
    gf = slide.shapes.add_table(nrow, ncol, Inches(L), Inches(T), Inches(W), Inches(height))
    t = gf.table
    if kind == "label_value":
        lab_w = min(1.7, W * 0.24)
        t.columns[0].width = Inches(lab_w)
        t.columns[1].width = Inches(W - lab_w)
        hdr_rows = 1 if header else 0
        for ri, pair in enumerate(data):
            _replace_text_keep_runs(t.cell(ri, 0).text_frame, pair[0])
            vtf = t.cell(ri, 1).text_frame
            _replace_text_keep_runs(vtf, pair[1])
            if ri >= hdr_rows:   # 헤더행(구분|내용)은 가운데, 값행은 모든 줄 좌측정렬 통일
                for p in vtf.paragraphs:
                    p.alignment = PP_ALIGN.LEFT
        # 헤더행=네이비+흰색Bold, 구분열=회색Bold (분할 청크마다 헤더 반복)
        style_table(gf, has_header=bool(header), label_cols=(0,),
                    header_fill=PALETTE["navy_dark"], label_fill=PALETTE["label_gray"])
    else:
        widths = _content_col_widths(data, ncol, W, font_pt)
        for c in range(ncol):
            t.columns[c].width = Inches(widths[c])
        for ri, row in enumerate(data):
            for ci in range(ncol):
                _replace_text_keep_runs(
                    t.cell(ri, ci).text_frame,
                    str((row[ci] if ci < len(row) else "") or ""))
        style_table(gf, has_header=bool(header), label_cols=(), header_fill=PALETTE["navy_dark"])
        # ★기업개요 4열(구분|내용|구분|내용)은 세로병합 금지 — 각 행이 독립(업종명이 주소 행으로
        #   번지는 버그 방지). 대신 우측 쌍이 빈 행(주소 등)은 값(col1)을 우측 끝까지 가로 병합(전폭).
        is_corp4 = (ncol == 4 and len(header) == 4
                    and "구분" in str(header[0]) and "구분" in str(header[2]))
        if not is_corp4:
            _merge_vertical_runs(t, data, ncol, header_rows=1 if header else 0)  # 반복값 세로병합
        _merge_total_rows(t, data, ncol)                                     # 합계행 가로병합
        if is_corp4:
            for ri in range(1 if header else 0, nrow):
                r = data[ri]
                c2 = str((r[2] if len(r) > 2 else "") or "").strip()
                c3 = str((r[3] if len(r) > 3 else "") or "").strip()
                if not c2 and not c3 and str((r[1] if len(r) > 1 else "") or "").strip():
                    try:
                        t.cell(ri, 1).merge(t.cell(ri, 3))   # 주소 등 전폭 값
                    except Exception:
                        pass
    _compact_grid(t, font_pt, row_h, kind=kind, has_header=bool(header))
    if header and hdr_fill_hex:   # 중첩표 헤더 = 밝은색 + 투명도
        _set_header_fill_alpha(t, hdr_fill_hex, hdr_alpha if hdr_alpha is not None else 35)
    hdr_rows = 1 if header else 0
    # ★표 안의 표 앵커 행은 글씨를 위에 두고 그 아래 grid를 얹으므로 위 정렬(나머지는 가운데)
    if anchor_rows:
        for li in anchor_rows:
            r = hdr_rows + li
            if r < nrow:
                t.cell(r, 1).vertical_anchor = MSO_ANCHOR.TOP
    # ★구분열(content 표) 밝은 회색 — 원본: '구분' 칸은 항상 F2F2F2.
    #   기업개요 4열(구분|내용|구분|내용)이면 양쪽 구분열(col0, col2) 모두 회색(사용자 지적).
    is_content = (kind != "label_value")
    if is_content and header:
        gubun_cols = [c for c in range(ncol) if c < len(header) and "구분" in str(header[c] or "")]
        for ri in range(hdr_rows, nrow):
            for gc in gubun_cols:
                cl = t.cell(ri, gc)
                cl.fill.solid()
                cl.fill.fore_color.rgb = PALETTE["label_gray"]
    # ★합계행 음영(원본 규칙, 사용자 재확인):
    #   - 표 안의 표(중첩 grid, hdr_fill_hex 지정) → 합계 색칠 '무조건 안 함'
    #   - 그 외 표: 표의 '맨 밑' 합계/총합계(마지막 청크의 최하단 합계행) 1개만 하늘 민트(3E95BE 65%).
    #     중간에 끼인 소계·합계·총합계는 전부 한 톤 어두운 회색(D9D9D9).
    is_nested = bool(hdr_fill_hex)
    if not is_nested:
        subtotal_rows, grandtotal_rows = _classify_total_rows(data, ncol)
        total_rows = subtotal_rows | grandtotal_rows
        bottom_total = max(total_rows) if total_rows else -1
        for ri in range(hdr_rows, nrow):
            if ri not in total_rows:
                continue
            if ri == bottom_total and is_last_chunk:
                for ci in range(ncol):
                    _set_one_cell_fill_alpha(t.cell(ri, ci), "3E95BE", 35)   # 맨 밑 합계 = 민트 65%
            else:
                for ci in range(ncol):
                    cl = t.cell(ri, ci)
                    cl.fill.solid()
                    cl.fill.fore_color.rgb = PALETTE["gray"]                 # 중간 소계/합계 = D9D9D9
    # ★원본 빨간 글씨 재현: '모든 동일텍스트 칸이 빨강이었을 때만' 칠함(과다 방지) + 빨간 테두리
    if red_counts:
        body_counts = Counter()
        for ri in range(hdr_rows, nrow):
            for ci in range(ncol):
                ct = t.cell(ri, ci).text.strip()
                if ct:
                    body_counts[ct] += 1
        for ri in range(hdr_rows, nrow):
            for ci in range(ncol):
                cell = t.cell(ri, ci)
                ct = cell.text.strip()
                if ct and red_counts.get(ct, 0) > 0 and red_counts[ct] >= body_counts[ct]:
                    _set_cell_red(cell)
    # ★원본에서 셀 외곽이 빨간 강조('본건' 행 등, 시세표) → 빨간 테두리 1pt 재현
    for ri in range(hdr_rows, nrow):
        c0 = str((data[ri][0] if ri < len(data) and data[ri] else "") or "").strip()
        if c0 in ("본건", "본 건"):
            for ci in range(ncol):
                _set_cell_red_border(t.cell(ri, ci))
    return height


def _place_images_row(slide, imgs, L, top, W, bottom, *, labels=None):
    """남은 본문 공간에 이미지들을 한 줄로 배치(가로 균등, 높이 맞춰 축소). 사용높이(in) 반환.
       labels: 각 이미지 위 작은 라벨(조감도/광역입지 등) 리스트(선택)."""
    imgs = [im for im in (imgs or []) if im.get("data")]
    if not imgs:
        return 0.0
    avail_h = bottom - top
    if avail_h < 0.6:
        return 0.0
    n = len(imgs)
    gap = 0.2
    cell_w = (W - gap * (n - 1)) / n
    lab_h = 0.22 if labels else 0.0
    x = L
    used = 0.0
    for i, im in enumerate(imgs):
        iw, ih = im.get("width", 1), im.get("height", 1)
        scale = min(cell_w / iw, (avail_h - lab_h) / ih)
        w_in, h_in = iw * scale, ih * scale
        if labels and i < len(labels) and labels[i]:
            _add_textbox(slide, x, top, cell_w, lab_h, labels[i],
                         size=9, bold=True, align=PP_ALIGN.CENTER, color=PALETTE["navy_dark"])
        px = x + (cell_w - w_in) / 2
        try:
            slide.shapes.add_picture(io.BytesIO(im["data"]), Inches(px), Inches(top + lab_h),
                                     Inches(w_in), Inches(h_in))
        except Exception:
            pass
        used = max(used, lab_h + h_in)
        x += cell_w + gap
    return used


def _img_labels_for(subtitle, n):
    """소제목으로 이미지 라벨 추정(조감도/광역입지/현장사진/위치도 등)."""
    s = subtitle or ""
    if "승인" in s or "공문" in s or "인허가" in s:
        base = ["사업계획승인서", "첨부"]
    elif "현장" in s:
        base = ["현장사진", "위치도"]
    elif "건축" in s or "조감" in s:
        base = ["조감도", "광역입지"]
    elif "입지" in s or "시세" in s or "분양" in s or "토지" in s:
        base = ["위치도", "조감도"]
    else:
        base = ["참고 이미지", "참고 이미지"]
    return [base[i] if i < len(base) else base[-1] for i in range(n)]


def _place_images_col(slide, imgs, L, top, W, bottom, *, labels=None):
    """오른쪽 열에 '표 박스(네이비 라벨 헤더 + 셀)'를 세로로 쌓고, 셀 위에 사진을 얹는다.
       (사진만 붙이는 게 아니라 표처럼 깔끔하게 — 사용자 지시)"""
    imgs = [im for im in (imgs or []) if im.get("data")]
    if not imgs:
        return 0.0
    avail_h = bottom - top
    if avail_h < 0.7:
        return 0.0
    n = len(imgs)
    gap = 0.18
    box_h = (avail_h - gap * (n - 1)) / n
    hdr_h = 0.28
    y = top
    for i, im in enumerate(imgs):
        lab = (labels[i] if labels and i < len(labels) else "참고 이미지")
        # 2행 표(헤더 + 본문 셀)
        gf = slide.shapes.add_table(2, 1, Inches(L), Inches(y), Inches(W), Inches(box_h))
        tb = gf.table
        tb.cell(0, 0).text = lab
        tb.rows[0].height = Inches(hdr_h)
        tb.rows[1].height = Inches(box_h - hdr_h)
        style_table(gf, has_header=True, label_cols=(), header_fill=PALETTE["navy_dark"])
        for p in tb.cell(0, 0).text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.size = Pt(9.5)
        # 본문 셀 위에 사진 얹기(셀 안쪽 여백 0.05)
        cell_w, cell_h = W - 0.10, (box_h - hdr_h) - 0.10
        iw, ih = im.get("width", 1), im.get("height", 1)
        scale = min(cell_w / iw, cell_h / ih)
        w_in, h_in = iw * scale, ih * scale
        px = L + (W - w_in) / 2
        py = y + hdr_h + 0.05 + (cell_h - h_in) / 2
        try:
            slide.shapes.add_picture(io.BytesIO(im["data"]), Inches(px), Inches(py),
                                     Inches(w_in), Inches(h_in))
        except Exception:
            pass
        y += box_h + gap
    return avail_h


_GAP = 0.22       # 표 간 간격
_LABEL_H = 0.34   # 표 미니라벨 높이


def build_structured_slide(prs, struct: dict, *, business_name: str = "",
                           section_label: str = None, subtitle: str = None,
                           images=None, red_texts=None, underline_texts=None):
    """LLM 구조화 결과 → 1개 이상의 본문 슬라이드. 내용이 길면 헤더 반복하며 "(i/n)"로 분할."""
    section_label = (section_label if section_label is not None
                     else struct.get("section_label") or "").strip()
    subtitle = (subtitle if subtitle is not None
                else struct.get("subtitle") or "").strip()
    intro = (struct.get("intro") or "").strip()
    bullets = [b for b in (struct.get("bullets") or []) if str(b).strip()]
    tables = struct.get("tables") or []
    # ★기업개요(시행사·시공사 기본정보) label_value → 4열(구분|내용|구분|내용)로(원본처럼)
    tables = [_corp_to_4col(t) if _is_corp_info(t) else t for t in tables]
    source = (struct.get("source") or "").strip()

    # ★표 안의 표: 앵커 라벨 → grid 파싱 (해당 label_value 행 내용칸에 grid를 얹음)
    nested_grids = struct.get("_nested_grids") or []
    nested_parsed = []   # (anchor_label, (kind,header,body,ncol), note)
    for item in nested_grids:
        anchor, gdef = item[0], item[1]
        note = item[2] if len(item) > 2 else ""
        gk, gh, gb, gn = _parse_tdef(gdef)
        if gn > 0 and gb:
            nested_parsed.append((anchor, (gk, gh, gb, gn), note))

    def _nested_for(label):
        for a, parsed, note in nested_parsed:
            if a and a in str(label):
                return parsed
        return None

    def _nested_note_for(label):
        for a, parsed, note in nested_parsed:
            if a and a in str(label):
                return note
        return ""

    red_set = set(red_texts or [])             # 산문(글상자) 부분 빨강용
    red_counts = Counter(red_texts or [])      # 표 셀 빨강용(모든 동일칸 빨강일 때만)
    ul_set = set(underline_texts or [])        # 산문(글상자) 부분 밑줄용
    # ★주1)2)3) 등 각주는 본문이 아니라 맨 아래로(원본처럼). 그 외 설명만 본문 불릿.
    def _is_note(b):
        s = str(b).lstrip()
        return s.startswith(("주", "*", "※")) or s[:2] in ("1)", "2)", "3)", "4)", "5)")
    notes = [b for b in bullets if _is_note(b)]
    body_bullets = [b for b in bullets if b not in notes]
    btext = "\n".join(f"• {b}" for b in body_bullets) if body_bullets else ""

    # ★출처 필드에 섞여 들어온 주N)/※ 각주 분리 — 출처 글상자는 '출처'만, 주N)는 표 밑 각주로
    def _looks_note(ln):
        s = ln.lstrip("(").lstrip()
        return s.startswith(("주", "*", "※")) and (
            s[:1] in ("*", "※") or (len(s) > 1 and s[1].isdigit()) or s.startswith("주 "))
    if source:
        _slines = [x.strip() for x in source.replace(" / ", "\n").split("\n") if x.strip()]
        _src_keep = [x for x in _slines if not _looks_note(x)]
        _src_notes = [x for x in _slines if _looks_note(x)]
        source = " ".join(_src_keep).strip()
        notes = notes + _src_notes
    notes_text = "\n".join(notes) if notes else ""

    # ── 본문 사진 배치 규칙(사람 제안서 실측) ──
    #   • 원본이 '표 안 사진'(조감도·광역위치도 등 라벨 이미지) → 표박스(우측, 1:1) = side_box
    #   • 원본이 '그냥 사진'(담보토지·인근시장 지도) → 표 위 전폭 바 이미지   = top_bare
    #   • 표 없는 부록(현장사진·승인서) → 표 없이 크게 중앙
    big_imgs = _big_images(images)
    has_tbl = any(_parse_tdef(t)[3] > 0 and _parse_tdef(t)[2] for t in tables)
    labeled_img = bool(big_imgs) and any(k in subtitle for k in ("건축", "입지", "조감", "위치"))
    side_box = bool(big_imgs) and has_tbl and labeled_img
    top_bare = bool(big_imgs) and has_tbl and not labeled_img
    _IMG_W, _IMG_GAP_LR = 4.0, 0.25
    _sidebox_tw = _TBL_W - _IMG_W - _IMG_GAP_LR
    tw = _sidebox_tw if side_box else _TBL_W
    img_col_L = _TBL_L + tw + _IMG_GAP_LR
    _IMG_TOP_H = 2.8 if top_bare else 0.0
    # ★side_box(조감도/광역입지) 이미지는 '첫 plan'(건축개요 사업명·대지위치 표) 옆에 둔다.
    #   → 용도별공급계획(다음 plan)은 전폭 유지(좁아져 이상해지던 문제 해결).
    img_plan_idx = 0

    # ── 표를 청크로 분할 + 라벨에 "(k/m)" 부착 (★우측상단 표식 아님, 표 라벨에 이어 적음) ──
    #   각 표마다 빈 슬라이드 1장 용량 기준으로 m등분 → "(k/m)"이 안정적으로 결정됨.
    blocks = []   # (label, kind, header, rows, ncol, fp, rh)  rh=스칼라(grid)|리스트(label_value)
    for tdef in tables:
        kind, header, body, ncol = _parse_tdef(tdef)
        if ncol == 0 or not body:
            continue
        fp = _grid_font(ncol)
        title = (tdef.get("title") or "").strip()
        has_h = bool(header)
        label_h = _LABEL_H if title else 0.0

        if kind == "label_value":
            # ★내용이 긴 약정문 → 행별 높이 가변. 행높이 합이 한 장 용량 넘으면 분할(헤더 반복).
            #   ★앵커 행(주요 대출조건·자금용도)은 grid 높이만큼 키워 그 안에 grid를 얹는다(표 안의 표).
            lab_w = min(1.7, tw * 0.24)
            val_w = tw - lab_w
            hdr_h = _rowh(fp)
            avail = _BODY_BOTTOM - _INTRO_T - label_h - hdr_h
            rhs_all = []
            for r in body:
                ng = _nested_for(r[0])
                if ng:
                    _, gh, gb, gn = ng
                    grid_h = ((1 if gh else 0) + len(gb)) * _rowh(_grid_font(gn)) + 0.10
                    # ★앵커 행 = (요약 글씨) + (grid) + (그 표 바로 밑 주N) 각주)
                    txt_h = (_est_text_height(r[1], val_w - 0.08, fp) + 0.04
                             if str(r[1]).strip() else 0.0)
                    nt = _nested_note_for(r[0])
                    note_h = (_est_text_height(nt, val_w - 0.10, 9) + 0.04) if nt.strip() else 0.0
                    rhs_all.append(grid_h + txt_h + note_h)
                else:
                    rhs_all.append(max(_est_text_height(r[0], lab_w - 0.08, fp),
                                       _est_text_height(r[1], val_w - 0.08, fp)) + 0.04)
            chunks, cr, crh, cacc = [], [], [], 0.0
            for r, h in zip(body, rhs_all):
                if cr and cacc + h > avail:
                    chunks.append((cr, crh)); cr, crh, cacc = [], [], 0.0
                cr.append(r); crh.append(h); cacc += h
            if cr:
                chunks.append((cr, crh))
            m = len(chunks)
            for k, (rows, rhs) in enumerate(chunks):
                lbl = (f"{title}({k + 1}/{m})" if (title and m > 1) else title)
                anchors = [(i, rows[i][0]) for i in range(len(rows)) if _nested_for(rows[i][0])]
                blocks.append((lbl, kind, header, rows, ncol, fp, rhs, anchors))
        else:
            rh = _rowh(fp)
            cap_full = max(1, int((_BODY_BOTTOM - _INTRO_T - label_h) / rh) - (1 if has_h else 0))
            m = max(1, -(-len(body) // cap_full))        # 필요한 분할 수(ceil)
            csize = -(-len(body) // m)                   # 균등 분할 크기
            for k in range(m):
                start = k * csize
                chunk = [list(r) for r in body[start:start + csize]]   # 복사(수정 위해)
                if not chunk:
                    continue
                # ★다음 장으로 넘어간 세로병합값: 첫 행의 빈 칸을 직전 비어있지 않은 값으로 재표기
                #   (A구역처럼 위에서 병합된 라벨이 빈칸으로 시작하지 않도록)
                if start > 0:
                    for c in range(ncol):
                        cell0 = str(chunk[0][c] if c < len(chunk[0]) else "").strip()
                        if cell0:
                            continue
                        for r in range(start - 1, -1, -1):
                            prev = str(body[r][c]) if c < len(body[r]) else ""
                            if prev.strip():
                                while len(chunk[0]) <= c:
                                    chunk[0].append("")
                                chunk[0][c] = prev
                                break
                lbl = (f"{title}({k + 1}/{m})" if (title and m > 1) else title)
                blocks.append((lbl, kind, header, chunk, ncol, fp, rh, None))

    # ── 블록을 슬라이드에 채움 (각 청크는 빈 슬라이드 1장에 반드시 들어감) ──
    #   ★제목의 내용(인트로)이 없으면 본문을 인트로 자리(_INTRO_T)부터 시작 = 전체적으로 위로.
    plans = []
    cur = {"intro": intro or None, "items": []}
    top = _INTRO_T
    if intro:
        top = _INTRO_T + _est_text_height(intro, _INTRO_W, 9) + 0.12
    top += _IMG_TOP_H   # 슬1 상단 전폭 사진 공간(top_bare)
    # ★본문 산문(bullets)·각주(주N)는 '표 아래'에 배치(원본: 글이 표 끝에) → 하단 공간 예약
    reserve_bottom = 0.0
    if btext:
        reserve_bottom += _est_text_height(btext, tw, 9) + 0.18
    if notes_text:
        reserve_bottom += _est_text_height(notes_text, _TBL_W, 9) + 0.10
    pack_bottom = _BODY_BOTTOM - reserve_bottom

    def flush():
        nonlocal cur, top
        plans.append(cur)
        cur = {"intro": None, "items": []}
        top = _INTRO_T          # 연속 슬라이드(인트로 없음)는 인트로 자리부터 = 위로

    for blk in blocks:
        lbl, kind, header, rows, ncol, fp, rh, _anchors = blk
        label_h = _LABEL_H if lbl else 0.0
        if isinstance(rh, (list, tuple)):
            need = label_h + (_rowh(fp) if header else 0) + sum(rh) + _GAP
        else:
            need = label_h + ((1 if header else 0) + len(rows)) * rh + _GAP
        if cur["items"] and top + need > pack_bottom:
            flush()
        cur["items"].append(blk)
        top += need
    flush()

    # ── 렌더 ──
    n = len(plans)
    for idx, plan in enumerate(plans):
        slide = clone_slide_layout(prs, "content", skip_graphic_frames=True)
        _fill_header(slide, section_label, subtitle, "")   # 소제목은 목차와 동일
        # ★plan별 표 폭: side_box 이미지가 들어가는 plan만 좁게, 나머지는 전폭
        if side_box:
            tw = _sidebox_tw if idx == img_plan_idx else _TBL_W
            img_col_L = _TBL_L + tw + _IMG_GAP_LR
        t = _INTRO_T
        if plan["intro"]:
            # ★제목의 내용 글상자 = Bold (사용자 지시), 원본 빨간 글씨 부분만 빨강
            _, h = _add_textbox(slide, _INTRO_L, _INTRO_T, _INTRO_W, 0.40, plan["intro"],
                                size=9, bold=True, red_set=red_set, ul_set=ul_set)
            t = _INTRO_T + h + 0.12
        # (본문 산문 bullets는 더 이상 표 위에 찍지 않음 — 표 아래로 이동, 아래 렌더 참조)
        # ★'그냥 사진'(지도 등)은 표 위에 전폭으로 (사람 제안서: 담보토지·인근시장)
        if idx == 0 and top_bare:
            _place_images_row(slide, big_imgs[:1], _TBL_L, t, _TBL_W, t + _IMG_TOP_H)
            t += _IMG_TOP_H + 0.12
        tbl_start = t   # 표가 시작되는 y(사진 옆배치 기준)
        for (lbl, kind, header, rows, ncol, fp, rh, anchors) in plan["items"]:
            if lbl:
                # ★'(단위 …)'는 표 우측 위에 Light 9pt 별도 글상자로(원본처럼), 라벨에선 분리
                unit = ""
                mi = lbl.find("(단위")
                if mi != -1:
                    me = lbl.find(")", mi)
                    if me != -1:
                        unit = lbl[mi:me + 1]
                        lbl = (lbl[:mi] + lbl[me + 1:]).strip()
                    else:
                        unit = lbl[mi:]; lbl = lbl[:mi].strip()
                _add_textbox(slide, _TBL_L, t, tw - 2.2, 0.28, lbl,
                             size=12, bold=True, color=PALETTE["navy_dark"])
                if unit:
                    _add_textbox(slide, _TBL_L + tw - 2.4, t + 0.06, 2.4, 0.20, unit,
                                 size=9, bold=False, align=PP_ALIGN.RIGHT, color=PALETTE["gray_text"])
                t += _LABEL_H
            tbl_top = t
            _anchor_li = {li for (li, _a) in anchors} if anchors else None
            # ★분할표(k/m)에서 '맨 밑 합계=민트'는 마지막 청크에만 적용 → is_last_chunk 판단
            _is_last_chunk = True
            if lbl and lbl.rstrip().endswith(")") and "(" in lbl:
                _tail = lbl[lbl.rfind("(") + 1:lbl.rfind(")")]
                if "/" in _tail:
                    _a, _b = (_tail.split("/") + [""])[:2]
                    if _a.strip().isdigit() and _b.strip().isdigit():
                        _is_last_chunk = int(_a) >= int(_b)
            used = _render_table_chunk(slide, kind, header, rows, ncol, _TBL_L, t, tw, fp, rh,
                                       red_counts=red_counts, anchor_rows=_anchor_li,
                                       is_last_chunk=_is_last_chunk)
            # ★표 안의 표: 앵커 행 내용칸 위에 grid를 얹음
            if anchors and isinstance(rh, (list, tuple)):
                lab_w = min(1.7, tw * 0.24)
                val_x = _TBL_L + lab_w
                val_w = tw - lab_w
                hh = _rowh(fp)
                for (li, alabel) in anchors:
                    ng = _nested_for(alabel)
                    if not ng:
                        continue
                    gk, gh, gb, gn = ng
                    # 앵커 행의 요약 글씨 높이만큼 내려서 grid를 얹음(글씨 위 + 표 아래)
                    atext = str(rows[li][1] if li < len(rows) and len(rows[li]) > 1 else "")
                    txt_h = (_est_text_height(atext, val_w - 0.08, fp) + 0.04) if atext.strip() else 0.0
                    row_y = tbl_top + hh + sum(rh[:li]) + txt_h
                    grid_h = ((1 if gh else 0) + len(gb)) * _rowh(_grid_font(gn)) + 0.10
                    try:
                        _render_table_chunk(slide, "grid", gh, gb, gn,
                                            val_x + 0.05, row_y + 0.05, val_w - 0.10,
                                            _grid_font(gn), _rowh(_grid_font(gn)), red_counts=red_counts,
                                            hdr_fill_hex="3E95BE", hdr_alpha=35)
                        # ★그 grid '바로 밑'에 주N) 각주(원본처럼)
                        gnote = _nested_note_for(alabel)
                        if gnote.strip():
                            _add_textbox(slide, val_x + 0.05, row_y + 0.05 + grid_h + 0.02,
                                         val_w - 0.10, 0.2, gnote, size=9,
                                         color=PALETTE["gray_text"], red_set=red_set, ul_set=ul_set)
                    except Exception as _e:
                        print(f"[nested grid] 실패: {_e}")
            t += used + _GAP
        # ★본문 사진 배치(top_bare는 위에서 처리).
        #   side_box(조감도/광역입지)=첫 plan(건축개요 표) 오른쪽 / 부록=마지막 plan 표없이 크게
        if side_box and idx == img_plan_idx and big_imgs:
            labels = _img_labels_for(subtitle, min(2, len(big_imgs)))
            # 원본이 '표 안 사진'(조감도·광역위치도) → 라벨 헤더 표박스(각 1:1), 표 오른쪽
            _place_images_col(slide, big_imgs[:2], img_col_L, tbl_start,
                              _IMG_W, _BODY_BOTTOM, labels=labels)
        elif idx == n - 1 and big_imgs and not top_bare and not side_box and not has_tbl:
            # 표 없는 부록(현장사진·승인서) → 표 없이 사진만 크게(위쪽부터)
            _place_images_row(slide, big_imgs[:2], _TBL_L, tbl_start,
                              _TBL_W, _BODY_BOTTOM)
        # ★본문 산문(설명 bullets) = 표 '아래'(원본: 글이 표 끝에). 본문 글상자=10pt 검정
        #   (표 바로 밑 주N)/출처 9pt보다 크게 — 사용자 지시). Light체.
        if idx == n - 1 and btext:
            bh = _est_text_height(btext, tw, 10)
            _, _bhh = _add_textbox(slide, _TBL_L, t + 0.06, tw, bh, btext,
                                   size=10, bold=False, color=RGBColor(0x00, 0x00, 0x00),
                                   red_set=red_set, ul_set=ul_set)
            t += _bhh + 0.10
        # ★주N) 각주 = 표(또는 표아래 산문) '바로 밑'(별도 글상자). 빨간 부분 살림.
        if idx == n - 1 and notes_text:
            nh = _est_text_height(notes_text, _TBL_W, 9)
            _add_textbox(slide, _TBL_L, t + 0.04, _TBL_W, nh, notes_text,
                         size=9, color=PALETTE["gray_text"], red_set=red_set, ul_set=ul_set)
        _add_combined_footer(slide, business_name)
        # ★출처는 항상 좌측 하단 고정(별도 글상자) — 원본에 있을 때만
        if idx == n - 1 and source:
            _add_textbox(slide, _SRC_L, _SRC_T, _SRC_W, 0.22,
                         f"* 출처 : {source}", size=9, color=PALETTE["gray_text"])
    return None


# ──────────────────────────────────────────────────────
# 투자구조도 — 도형(상자/선/글상자)으로 직접 그림 (이미지 캡처 아님)
# ──────────────────────────────────────────────────────
def _set_line_arrow(conn, *, color="595959", width_pt=1.25, end=True, begin=False):
    """커넥터 선 색·두께·화살표 머리 설정(XML)."""
    ln = conn.line._get_or_add_ln()
    ln.set("w", str(int(width_pt * 12700)))
    # 색
    for tag in ("a:solidFill", "a:noFill"):
        for el in ln.findall(qn(tag)):
            ln.remove(el)
    sf = ln.makeelement(qn("a:solidFill"), {})
    sf.append(sf.makeelement(qn("a:srgbClr"), {"val": color}))
    ln.insert(0, sf)
    if begin:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if end:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))


def _dbox(slide, l, t, w, header, bodies, *, hdr_cm=0.9, body_cm=1.7, header_fill=None,
          body_fill=None, body_white=False):
    """다이어그램 상자 = 2~N행 1열 표(헤더 색 + 본문). graphicFrame 반환.
       ★고정 사양: 헤더행 0.9cm, 본문행 1.7cm(차주 기준), 글꼴 전부 Bold 10.5pt.
       body_fill 지정 시 본문 셀 색 변경(대주단 트랜치=진남색+흰글자 등)."""
    if isinstance(bodies, str):
        bodies = [bodies] if bodies else []
    bodies = [str(b) for b in bodies if str(b).strip()]
    nrow = 1 + len(bodies)
    total_h = Cm(hdr_cm + body_cm * len(bodies))
    gf = slide.shapes.add_table(nrow, 1, Inches(l), Inches(t), Inches(w), total_h)
    tb = gf.table
    # ★v22처럼 표 자체(tblPr)에 드롭섀도 추가(blurRad 50800/dist 38100/dir 2700000/검정 alpha 40%)
    _tblPr = tb._tbl.find(qn("a:tblPr"))
    if _tblPr is not None:
        for _e in _tblPr.findall(qn("a:effectLst")):
            _tblPr.remove(_e)
        _eff = _tblPr.makeelement(qn("a:effectLst"), {})
        _shd = _eff.makeelement(qn("a:outerShdw"),
                                {"blurRad": "50800", "dist": "38100", "dir": "2700000",
                                 "algn": "tl", "rotWithShape": "0"})
        _c = _shd.makeelement(qn("a:prstClr"), {"val": "black"})
        _c.append(_c.makeelement(qn("a:alpha"), {"val": "40000"}))
        _shd.append(_c)
        _eff.append(_shd)
        _tblPr.insert(0, _eff)
    tb.cell(0, 0).text = header
    for i, b in enumerate(bodies):
        tb.cell(1 + i, 0).text = b
    # ★헤더=지정색(기본 네이비), 본문=회색 D9D9D9(기본) 또는 body_fill
    style_table(gf, has_header=True, label_cols=(),
                header_fill=header_fill or PALETTE["navy_dark"],
                value_fill=body_fill or PALETTE["gray"])
    # ★행 높이는 style_table(헤더 0.6cm로 덮어씀) 뒤에 다시 강제 — 헤더 0.9cm / 본문 1.7cm
    tb.rows[0].height = Cm(hdr_cm)
    for i in range(1, nrow):
        tb.rows[i].height = Cm(body_cm)
    for ri in range(nrow):
        cell = tb.cell(ri, 0)
        cell.margin_top = cell.margin_bottom = Inches(0.01)
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.size = Pt(13 if ri == 0 else 12)   # ★표 박스 글씨 헤더13/본문12pt(사용자 지시)
                r.font.name = _FONT_BOLD                    # ★전부 Bold 고정
                if ri >= 1 and body_white:
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # ★v22처럼 4변 흰색 테두리 2.25pt(28575) — 박스 사이 흰 간격
        tcPr = cell._tc.get_or_add_tcPr()
        for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
            for el in tcPr.findall(qn(tag)):
                tcPr.remove(el)
            ln = tcPr.makeelement(qn(tag), {"w": "28575", "cap": "flat", "cmpd": "sng", "algn": "ctr"})
            sf = ln.makeelement(qn("a:solidFill"), {})
            sf.append(sf.makeelement(qn("a:srgbClr"), {"val": "FFFFFF"}))
            ln.append(sf)
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
            tcPr.insert(i, ln)
    return gf


def _dlabel(slide, l, t, w, text):
    """선 위 설명 글상자(가운데, Bold 10.5pt). ★글상자 10.5pt + 전부 Bold + 검정(사용자 지시)."""
    _add_textbox(slide, l, t, w, 0.20, text, size=10.5, bold=True, align=PP_ALIGN.CENTER,
                 color=RGBColor(0x00, 0x00, 0x00))


def build_invest_diagram_slide(prs, data: dict, *, section_label, subtitle,
                               intro="", business_name=""):
    """투자구조도(2.1) — 참여기관을 '표 박스(네이비 헤더+본문)'와 연결선으로 그림(v22 방식)."""
    ent = (data or {}).get("entities", {}) or {}
    tranches = (data or {}).get("tranches", []) or []

    def _name(key, default):
        v = str(ent.get(key, "") or "").strip() or default
        for sep in ("—", "–", " - ", ":"):     # '신탁사 — 한국투자…' → '한국투자…'
            if sep in v:
                v = v.split(sep)[-1].strip()
        return v

    slide = clone_slide_layout(prs, "content", skip_graphic_frames=True)
    _fill_header(slide, section_label, subtitle, "")
    if intro and intro.strip():
        _add_textbox(slide, _INTRO_L, _INTRO_T, _INTRO_W, 0.40, intro.strip(), size=9, bold=True)

    # ★페이지 상·하단 가로 구분선(예시처럼)
    for ly in (1.50, 7.05):
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                          Inches(0.43), Inches(ly), Inches(10.40), Inches(ly))
        _set_line_arrow(line, color="BFBFBF", width_pt=1.0, end=False)

    # ★v22 3색 헤더: 차주=진남색, 신탁사·시공사·SPC·사채권자=중간파랑, 대주단=적갈색
    navy = PALETTE["navy_dark"]            # 08377C
    blue = RGBColor(0x00, 0x63, 0xA1)      # 0063A1
    maroon = RGBColor(0x8C, 0x4A, 0x59)    # 8C4A59

    # ── 표 박스 배치 (in) — 크게(보기 편하게). 신탁사·차주 같은 행, 시공사 차주 아래, 대주 우측 ──
    trustee = _dbox(slide, 0.70, 2.75, 2.05, "신탁사", _name("trustee", "한국투자부동산신탁㈜"),
                    header_fill=blue)
    borrower = _dbox(slide, 4.20, 2.75, 2.05, "차주", _name("borrower", "㈜루시드"),
                     header_fill=navy)
    constructor = _dbox(slide, 4.20, 4.70, 2.05, "시공사", _name("constructor", "롯데건설㈜"),
                        header_fill=blue)

    # 대주단(트랜치 다행) — 우측. v22 색: 적갈색 헤더 + 회색 본문
    tr_lines = []
    for tr in tranches:
        nm = tr.get("name", "") if isinstance(tr, dict) else str(tr)
        amt = tr.get("amount", "") if isinstance(tr, dict) else ""
        tr_lines.append((nm + (f" {amt}" if amt else "")).strip())
    if not tr_lines:
        tr_lines = ["Tr.A", "Tr.B", "Tr.C"]
    lenders = _dbox(slide, 7.90, 2.45, 2.20, "본건 PF 대주단", tr_lines,
                    hdr_cm=0.9, body_cm=0.95, header_fill=maroon)
    # ★원본 구조도는 박스 4개(신탁사·차주·시공사·대주단)만 — SPC/사채권자 없음

    def cx(sh, fx=0.5):
        return (sh.left + int(sh.width * fx)) / 914400

    def cy(sh, fy=0.5):
        return (sh.top + int(sh.height * fy)) / 914400

    def conn(x1, y1, x2, y2, **kw):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                       Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        _set_line_arrow(c, **kw)
        return c

    # ── 연결선 + 라벨 (원본 라벨 그대로) ──
    # 신탁사 → 차주 (사업부지 담보신탁 계약 체결, 수평)
    conn(cx(trustee, 1.0), cy(trustee), cx(borrower, 0.0), cy(borrower))
    _dlabel(slide, 2.80, 3.12, 1.55, "사업부지 담보신탁 계약 체결")
    # 차주 → 시공사 (공동사업약정 체결, 수직)
    conn(cx(borrower), cy(borrower, 1.0), cx(constructor), cy(constructor, 0.0), begin=True)
    _dlabel(slide, 6.30, 4.10, 1.5, "공동사업약정 체결")
    # 차주 ↔ 대주단 (대출실행/원리금상환, 수평)
    conn(cx(borrower, 1.0), cy(borrower), cx(lenders, 0.0), cy(lenders, 0.55), begin=True)
    _dlabel(slide, 6.35, 3.02, 1.55, "대출실행 / 원리금상환")
    # 시공사 → 대주단 (터미널 부지 물상보증, 대각)
    conn(cx(constructor, 1.0), cy(constructor), cx(lenders, 0.0), cy(lenders, 0.9))
    _dlabel(slide, 6.35, 4.95, 1.55, "터미널 부지 물상보증")
    # 신탁사 → 대주단 (순위별 우선수익권 제공, 상단 아치 3구간)
    ax = 2.30
    conn(cx(trustee), cy(trustee, 0.0), cx(trustee), ax)            # 위로
    conn(cx(trustee), ax, cx(lenders), ax)                          # 가로질러
    conn(cx(lenders), ax, cx(lenders), cy(lenders, 0.0), end=True)  # 대주단으로 내려
    _dlabel(slide, (cx(trustee) + cx(lenders)) / 2 - 1.0, ax - 0.22, 2.0, "순위별 우선수익권 제공")

    _add_combined_footer(slide, business_name)
    return None


# ──────────────────────────────────────────────────────
# 투자구조도 — 원본 PDF의 구조도 영역을 이미지로 캡처해 삽입(딜마다 달라 자동그리기 대신 캡처)
# ──────────────────────────────────────────────────────
def render_structure_image(pdf_path, dpi_scale=3):
    """원본 PDF에서 '구조도'가 있는 페이지의 다이어그램 영역을 PNG로 렌더 → {data,w,h} 또는 None."""
    try:
        import fitz
    except Exception:
        return None
    doc = fitz.open(pdf_path)
    try:
        target = None
        for i in range(len(doc)):
            if "구조도" in (doc[i].get_text("text") or ""):
                target = i
                break
        if target is None:
            return None
        page = doc[target]
        # '구조도' 라벨 y 위치
        ky = None
        for b in page.get_text("dict").get("blocks", []):
            if b.get("type") != 0:
                continue
            t = "".join(s["text"] for ln in b.get("lines", []) for s in ln.get("spans", []))
            if "구조도" in t:
                ky = b["bbox"][1]
                break
        # 라벨 아래 벡터 드로잉들의 합집합 bbox
        xs, ys = [], []
        for d in page.get_drawings():
            r = d["rect"]
            if r.y0 >= (ky - 4 if ky else 0) and r.width < page.rect.width and r.height < page.rect.height:
                xs += [r.x0, r.x1]
                ys += [r.y0, r.y1]
        if not xs:
            return None
        pad = 6
        clip = fitz.Rect(max(0, min(xs) - pad), max(0, min(ys) - pad),
                         min(page.rect.width, max(xs) + pad),
                         min(page.rect.height, max(ys) + pad))
        if clip.width < 40 or clip.height < 40:
            return None
        pix = page.get_pixmap(clip=clip, matrix=fitz.Matrix(dpi_scale, dpi_scale))
        return {"data": pix.tobytes("png"), "w": pix.width, "h": pix.height}
    finally:
        doc.close()


def build_invest_image_slide(prs, *, section_label, subtitle, intro, img,
                             business_name=""):
    """투자구조도 슬라이드 — 섹션라벨/소제목/인트로 + 캡처한 구조도 이미지(본문 중앙 배치)."""
    slide = clone_slide_layout(prs, "content", skip_graphic_frames=True)
    _fill_header(slide, section_label, subtitle, "")
    top = _BODY_TOP_START
    if intro and intro.strip():
        _, h = _add_textbox(slide, _INTRO_L, _INTRO_T, _INTRO_W, 0.40, intro.strip(), size=9)
        top = max(top, _INTRO_T + h + 0.12)
    if img and img.get("data"):
        avail_h = _BODY_BOTTOM - top - 0.1
        avail_w = _TBL_W
        scale = min(avail_w / img["w"], avail_h / img["h"])
        iw, ih = img["w"] * scale, img["h"] * scale
        px = _TBL_L + (avail_w - iw) / 2
        slide.shapes.add_picture(io.BytesIO(img["data"]), Inches(px), Inches(top),
                                 Inches(iw), Inches(ih))
    _add_combined_footer(slide, business_name)
    return slide


# ──────────────────────────────────────────────────────
# 푸터 — 우측에 "사업명ㅣ쪽번호" 합쳐서 (대전 방식). 쪽번호는 슬라이드번호 필드.
# ──────────────────────────────────────────────────────
def _add_combined_footer(slide, business_name):
    # 기존 템플릿 푸터(사업명 자리표시자 + '슬라이드 번호' 도형) 제거 → 중복 방지
    for sh in list(slide.shapes):
        nm = sh.name or ""
        is_pagenum = "슬라이드 번호" in nm
        is_biz_footer = (sh.has_text_frame and (sh.left or 0) / 360000 < 1.0
                         and (sh.top or 0) / 360000 > 17.5)
        if is_pagenum or is_biz_footer:
            try:
                sh._element.getparent().remove(sh._element)
            except Exception:
                pass

    tb = slide.shapes.add_textbox(Inches(7.18), Inches(7.19), Inches(3.32), Inches(0.25))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT

    run = p.add_run()
    run.text = f"{business_name or ''}｜"   # 전각 세로막대 ｜ (대전 원본 방식)
    run.font.name = _FONT_LIGHT
    run.font.size = Pt(8)
    run.font.color.rgb = PALETTE["gray_text"]

    # 슬라이드 번호 필드 (자동 쪽번호)
    fld = p._p.makeelement(qn('a:fld'), {
        'id': '{B7E3A1C2-0000-4000-9000-000000000001}', 'type': 'slidenum'})
    rPr = fld.makeelement(qn('a:rPr'), {'lang': 'ko-KR', 'sz': '800'})
    sf = rPr.makeelement(qn('a:solidFill'), {})
    clr = sf.makeelement(qn('a:srgbClr'), {'val': '808080'})
    sf.append(clr)
    rPr.append(sf)
    latin = rPr.makeelement(qn('a:latin'), {'typeface': _FONT_LIGHT})
    ea = rPr.makeelement(qn('a:ea'), {'typeface': _FONT_LIGHT})
    rPr.append(latin)
    rPr.append(ea)
    t_el = fld.makeelement(qn('a:t'), {})
    t_el.text = "1"
    fld.append(rPr)
    fld.append(t_el)
    p._p.append(fld)
    return tb
