import streamlit as st

# 추출 함수 및 UI 컴포넌트 불러오기
from extractors import extract_from_pdf, extract_from_docx, detect_business_name, get_file_type
from ui_components import render_stepper, render_image_gallery, render_text_preview

# PPT 생성 모듈 불러오기
import os
import re
from datetime import datetime
from modules.page_builders import (
    make_output_filename,
    build_full_presentation,
    build_preview_presentation,
)
from modules.content_parser import parse_document_from_bytes

# ─────────────────────────────────────────────
# [페이지 기본 설정]
# - page_title : 브라우저 탭에 표시되는 제목 (여기를 수정하면 탭 제목이 바뀝니다)
# - layout     : "wide" = 화면 전체 너비 사용 / "centered" = 가운데 정렬 좁은 화면
# - initial_sidebar_state : "collapsed" = 사이드바 기본 접힘 / "expanded" = 기본 펼침
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="레인필드투자자문 IM",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# [접근 코드 불러오기]
# - secrets.toml 파일에서 ACCESS_CODE 값을 읽어옵니다
# - 파일이 없거나 키가 없으면 기본값 'rainfield2026' 사용
# - 실제 운영 시에는 secrets.toml 에서만 관리하세요
# ─────────────────────────────────────────────
try:
    CORRECT_CODE = st.secrets["ACCESS_CODE"]
except Exception:
    CORRECT_CODE = "rainfield2026"  # ← 여기를 수정하면 기본 접근 코드가 바뀝니다

# ─────────────────────────────────────────────
# [세션 상태 초기화]
# - st.session_state : 페이지가 새로고침돼도 값이 유지되는 저장공간
# - logged_in : 로그인 여부 (True = 로그인됨, False = 로그아웃 상태)
# ─────────────────────────────────────────────
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# ─────────────────────────────────────────────
# [변환 작업 관련 세션 상태 초기화]
# 새로고침해도 데이터가 유지됩니다
# ─────────────────────────────────────────────
if "current_step" not in st.session_state:
    st.session_state.current_step = 1       # 현재 진행 단계 (1~9)
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None   # 업로드된 파일 객체
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None  # 추출된 텍스트·이미지 데이터
if "business_name" not in st.session_state:
    st.session_state.business_name = ""     # 자동 감지된 사업명

# 3단계 관련 세션 상태
if "toc_count" not in st.session_state:
    st.session_state.toc_count = 4          # 목차 개수 (4 또는 5)
if "month_en" not in st.session_state:
    st.session_state.month_en = datetime.now().strftime("%B")  # 현재 월 영문
if "year" not in st.session_state:
    st.session_state.year = str(datetime.now().year)           # 현재 연도
if "cover_image_index" not in st.session_state:
    st.session_state.cover_image_index = 0  # 표지에 쓸 이미지 번호
if "cover_image_bytes" not in st.session_state:
    st.session_state.cover_image_bytes = None  # 선택된 표지 이미지 bytes
if "parsed_pages" not in st.session_state:
    st.session_state.parsed_pages = []         # content_parser 결과

# 섹션 이미지 관련 세션 상태 — 원형 슬롯 3개 (4개 섹션 divider 공통 적용)
if "section_img_idx_list" not in st.session_state:
    st.session_state.section_img_idx_list = [0, 0, 0]
if "section_img_bytes_list" not in st.session_state:
    st.session_state.section_img_bytes_list = [None, None, None]

# 목차 이미지 관련 세션 상태 — 원형 슬롯 1개
if "toc_img_idx" not in st.session_state:
    st.session_state.toc_img_idx = 0
if "toc_img_bytes" not in st.session_state:
    st.session_state.toc_img_bytes = None


# ─────────────────────────────────────────────
# [로그인 화면 함수]
# ─────────────────────────────────────────────
def show_login():
    # 화면 가운데 정렬을 위해 3개 컬럼 중 가운데만 사용
    col1, col2, col3 = st.columns([1, 1.5, 1])

    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)

        # 로그인 화면 제목 (여기를 수정하면 로그인 화면 제목이 바뀝니다)
        st.markdown(
            "<h2 style='text-align:center;'>🌧 레인필드투자자문 IM</h2>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<p style='text-align:center; color:gray;'>접근 코드를 입력하세요</p>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # st.form : 폼 안에서 Enter 키를 누르면 버튼과 동일하게 동작합니다
        with st.form("login_form"):
            # 접근 코드 입력 필드
            # type="password" : 입력값이 ●●● 로 가려집니다
            code_input = st.text_input(
                label="접근 코드",
                type="password",
                placeholder="접근 코드 입력",
                label_visibility="collapsed",
            )

            # 로그인 버튼 (여기를 수정하면 버튼 텍스트가 바뀝니다)
            login_btn = st.form_submit_button("로그인", use_container_width=True, type="primary")

        if login_btn:
            if code_input == CORRECT_CODE:
                # 올바른 코드 → 로그인 성공
                st.session_state.logged_in = True
                st.rerun()  # 화면을 메인으로 전환
            else:
                # 잘못된 코드 → 빨간 오류 메시지
                st.error("❌ 접근 코드가 올바르지 않습니다. 다시 확인해주세요.")


# ─────────────────────────────────────────────
# [1단계: 파일 업로드 화면]
# ─────────────────────────────────────────────
def show_step1():
    st.markdown("## 1단계. 파일 업로드")
    st.caption("증권사·은행에서 받은 IM 자료를 업로드하세요. PDF 또는 Word(.docx) 파일만 지원됩니다.")
    st.markdown("")

    # 파일 업로더 (여기를 수정하면 허용 파일 형식이 바뀝니다)
    uploaded = st.file_uploader(
        label="파일을 여기에 끌어다 놓거나 클릭해서 선택하세요",
        type=["pdf", "docx"],
        accept_multiple_files=False,
        key="file_uploader_widget",
    )

    if uploaded is not None:
        # 파일 정보 표시
        file_size_mb = uploaded.size / (1024 * 1024)
        file_type_label = "PDF" if uploaded.name.lower().endswith(".pdf") else "Word(.docx)"

        # 50MB 이상 경고 (여기를 수정하면 경고 기준 크기가 바뀝니다)
        if file_size_mb >= 50:
            st.warning("⚠️ 파일이 큽니다. 추출에 시간이 걸릴 수 있습니다.")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("파일명", uploaded.name)
        with col2:
            st.metric("파일 크기", f"{file_size_mb:.2f} MB")
        with col3:
            st.metric("파일 종류", file_type_label)

        st.markdown("")

        # 추출 시작 버튼
        if st.button("🔍 추출 시작", type="primary", use_container_width=False):
            file_bytes = uploaded.read()
            file_type  = get_file_type(uploaded.name)

            try:
                with st.spinner("파일에서 텍스트와 이미지를 추출하는 중입니다..."):
                    if file_type == "pdf":
                        result = extract_from_pdf(file_bytes)
                    else:
                        result = extract_from_docx(file_bytes)

                # 추출 결과 저장
                st.session_state.extracted_data = result
                st.session_state.uploaded_file  = uploaded.name

                # 사업명 자동 감지
                detected = detect_business_name(result["pages_text"])
                st.session_state.business_name = detected

                # 슬라이드 구조 파싱 (content_parser)
                try:
                    st.session_state.parsed_pages = parse_document_from_bytes(
                        file_bytes, uploaded.name
                    )
                except Exception:
                    st.session_state.parsed_pages = []

                # 2단계로 이동
                st.session_state.current_step = 2
                st.rerun()

            except ValueError as e:
                # 비밀번호 걸린 PDF
                st.error(f"🔒 {e}")
            except RuntimeError as e:
                # 손상된 파일
                st.error(f"❌ {e}")
            except Exception as e:
                st.error(f"❌ 예상치 못한 오류가 발생했습니다: {e}")


# ─────────────────────────────────────────────
# [2단계: 추출 결과 확인 화면]
# ─────────────────────────────────────────────
def show_step2():
    data = st.session_state.extracted_data

    # 혹시 데이터가 없으면 1단계로 돌려보냄 (비정상 접근 방어)
    if data is None:
        st.warning("추출 데이터가 없습니다. 파일을 다시 업로드해주세요.")
        if st.button("← 1단계로 돌아가기"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.markdown("## 2단계. 추출 결과 확인")
    st.caption("추출된 내용을 확인해주세요. 사업명이 잘못 인식된 경우 수정할 수 있습니다.")
    st.markdown("")

    # ────────────────────────────────────────
    # (A) 핵심 정보 자동 감지
    # ────────────────────────────────────────
    st.markdown("### 📌 핵심 정보 자동 감지")

    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.metric("총 페이지 수", f"{data['total_pages']}페이지")
    with col_info2:
        st.metric("추출된 이미지 수", f"{len(data['images'])}개")
    with col_info3:
        st.metric("파일 형식", data["file_type"].upper())

    st.markdown("")

    # 사업명 입력란 — 자동 감지 값이 기본으로 채워지며 수정 가능
    # 여기를 수정하면 사업명 입력란 안내 문구가 바뀝니다
    if not st.session_state.business_name:
        st.warning("⚠️ 사업명을 자동으로 인식하지 못했습니다. 직접 입력해주세요.")

    business_input = st.text_input(
        label="사업명",
        value=st.session_state.business_name,
        placeholder="예) 천안 부성2지구 도시개발사업",
        help="자동 감지된 사업명입니다. 잘못된 경우 직접 수정해주세요.",
    )
    # 입력값을 세션에 즉시 반영
    st.session_state.business_name = business_input

    # 추출 중 경고 메시지 표시
    for w in data.get("warnings", []):
        st.warning(f"⚠️ {w}")

    st.markdown("---")

    # ────────────────────────────────────────
    # (B) 페이지별 텍스트 미리보기
    # ────────────────────────────────────────
    render_text_preview(data["pages_text"])

    st.markdown("---")

    # ────────────────────────────────────────
    # (C) 이미지 갤러리
    # ────────────────────────────────────────
    render_image_gallery(data["images"])

    st.markdown("---")

    # ────────────────────────────────────────
    # (D) 하단 네비게이션 버튼
    # ────────────────────────────────────────
    col_prev, col_space, col_next = st.columns([3, 2, 3])

    with col_prev:
        if st.button("← 이전 단계  (파일 다시 업로드)", use_container_width=True):
            # 추출 데이터 초기화 후 1단계로 복귀
            st.session_state.current_step   = 1
            st.session_state.extracted_data = None
            st.session_state.uploaded_file  = None
            st.session_state.business_name  = ""
            # 전체 펼치기/접기 상태도 초기화
            if "text_expanded" in st.session_state:
                del st.session_state["text_expanded"]
            st.rerun()

    with col_next:
        # 사업명이 없으면 다음 단계 버튼 비활성화
        next_disabled = not bool(st.session_state.business_name.strip())

        if next_disabled:
            st.warning("사업명을 입력해주세요.")

        if st.button(
            "다음 단계 →",
            use_container_width=True,
            type="primary",
            disabled=next_disabled,
        ):
            st.session_state.current_step = 3
            st.rerun()


# ─────────────────────────────────────────────
# [섹션 divider 슬롯 위치 미니어처 헬퍼]
# ─────────────────────────────────────────────
@st.cache_data
def _make_divider_miniature(w_px: int = 540, h_px: int = 374) -> bytes:
    """섹션 divider 슬라이드 레이아웃을 단순화한 PNG 미니어처를 반환합니다.
    ①②③ 표시로 각 원형 슬롯 위치를 직관적으로 안내합니다."""
    import io as _mio
    from PIL import Image as _Img, ImageDraw as _Draw, ImageFont as _Font

    SLIDE_W = 27.52
    SLIDE_H = 19.05
    sx, sy = w_px / SLIDE_W, h_px / SLIDE_H

    img  = _Img.new("RGB", (w_px, h_px), (248, 248, 248))
    draw = _Draw.Draw(img)
    draw.rectangle([0, 0, w_px - 1, h_px - 1], outline=(200, 200, 200), width=2)

    # 텍스트 영역 힌트 (섹션 번호+제목 위치)
    draw.rectangle(
        [int(0.5 * sx), int(7.5 * sy), int(13 * sx), int(10.5 * sy)],
        fill=(225, 225, 225), outline=(190, 190, 190),
    )

    # 3개 원형 슬롯 (outer oval 위치 기준) — _DIV_OVAL_PAIRS와 동일 순서
    _OVALS  = [(16.3273, 0.7346, 9.70, 9.70),
               (17.3779, 8.2234, 8.70, 8.70),
               (13.5919, 6.1749, 6.80, 6.80)]
    _LABELS = ["①", "②", "③"]
    _GREEN  = (146, 208, 80)

    try:
        _fnt_path = os.path.join(os.path.dirname(__file__), "fonts", "PEOPLEFONTB.TTF")
        _base_fnt = _Font.truetype(_fnt_path, size=28)
    except Exception:
        _base_fnt = _Font.load_default()

    for (ol, ot, ow, oh), label in zip(_OVALS, _LABELS):
        cx = int((ol + ow / 2) * sx)
        cy = int((ot + oh / 2) * sy)
        rx = int(ow / 2 * sx)
        ry = int(oh / 2 * sy)
        draw.ellipse([cx - rx, cy - ry, cx + rx, cy + ry],
                     fill=(220, 220, 220), outline=_GREEN, width=5)
        draw.text((cx, cy), label, font=_base_fnt, fill=(50, 50, 50), anchor="mm")

    buf = _mio.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
# [목차 TOC 슬롯 위치 미니어처 헬퍼]
# ─────────────────────────────────────────────
@st.cache_data
def _make_toc_miniature(w_px: int = 540, h_px: int = 374) -> bytes:
    """목차 슬라이드 레이아웃을 단순화한 PNG 미니어처를 반환합니다.
    ① 표시로 원형 슬롯 위치를 직관적으로 안내합니다."""
    import io as _mio
    from PIL import Image as _Img, ImageDraw as _Draw, ImageFont as _Font

    SLIDE_W = 27.52
    SLIDE_H = 19.05
    sx, sy = w_px / SLIDE_W, h_px / SLIDE_H

    img  = _Img.new("RGB", (w_px, h_px), (248, 248, 248))
    draw = _Draw.Draw(img)
    draw.rectangle([0, 0, w_px - 1, h_px - 1], outline=(200, 200, 200), width=2)

    # 오른쪽 목차 항목 힌트 (4개 행)
    for row in range(4):
        y_top = int((2.0 + row * 4.1) * sy)
        draw.rectangle(
            [int(7.5 * sx), y_top, int(26.5 * sx), int(y_top + 3.2 * sy)],
            fill=(225, 225, 225), outline=(190, 190, 190),
        )

    # 1개 원형 슬롯 — TOC oval 위치 (left=0.84, top=6.95, w=5.26, h=5.30)
    _GREEN = (146, 208, 80)
    ol, ot, ow, oh = 0.8409, 6.9520, 5.2600, 5.3000
    cx = int((ol + ow / 2) * sx)
    cy = int((ot + oh / 2) * sy)
    rx = int(ow / 2 * sx)
    ry = int(oh / 2 * sy)

    try:
        _fnt_path = os.path.join(os.path.dirname(__file__), "fonts", "PEOPLEFONTB.TTF")
        _base_fnt = _Font.truetype(_fnt_path, size=28)
    except Exception:
        _base_fnt = _Font.load_default()

    draw.ellipse([cx - rx, cy - ry, cx + rx, cy + ry],
                 fill=(220, 220, 220), outline=_GREEN, width=5)
    draw.text((cx, cy), "①", font=_base_fnt, fill=(50, 50, 50), anchor="mm")

    buf = _mio.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
# [목차 개수 자동 감지 헬퍼]
# ─────────────────────────────────────────────
def _detect_toc_count(full_text: str) -> int:
    """
    추출 텍스트에서 '01', '02' 같은 섹션 번호 패턴을 세어 목차 개수를 추정합니다.
    찾지 못하면 기본값 4를 반환합니다.
    여기를 수정하면 목차 감지 패턴이 바뀝니다.
    """
    matches = re.findall(r'\b0[1-9]\b', full_text)
    unique  = len(set(matches))
    if unique >= 5:
        return 5
    elif unique >= 3:
        return 4
    return 4  # 기본값


# ─────────────────────────────────────────────
# [3단계: 레이아웃 및 표지 미리 생성 화면]
# ─────────────────────────────────────────────
def show_step3():
    data = st.session_state.extracted_data
    if data is None:
        st.warning("추출 데이터가 없습니다. 1단계로 돌아가주세요.")
        if st.button("← 1단계로"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.markdown("## 3단계. 레이아웃 및 표지 미리 생성")
    st.caption("레이아웃을 선택하고 날짜·표지 이미지를 지정한 뒤 표지를 미리 생성해보세요.")
    st.markdown("")

    # ────────────────────────────────────────
    # (A) 레이아웃 자동 추천
    # ────────────────────────────────────────
    st.markdown("### 📐 레이아웃 자동 추천")

    detected_toc = _detect_toc_count(data.get("full_text", ""))
    st.info(f"추출된 텍스트 분석 결과: **목차 {detected_toc}개 형식** 을 추천합니다.")

    toc_choice = st.radio(
        "목차 개수를 선택하세요",
        options=[4, 5],
        index=0 if detected_toc <= 4 else 1,
        format_func=lambda x: f"목차 {x}개 형식",
        horizontal=True,
        key="toc_radio",
    )
    st.session_state.toc_count = toc_choice

    # Task 4: 선택한 목차 수와 실제 추출 섹션 수가 다를 때 경고
    _parsed = st.session_state.parsed_pages
    if _parsed:
        _sec_titles = list(dict.fromkeys(
            p.get("section_title", "").strip() for p in _parsed
            if p.get("section_title", "").strip()
        ))
        _actual_count = len(_sec_titles)
        if _actual_count > 0 and toc_choice > _actual_count:
            st.warning(
                f"⚠️ 목차 {toc_choice}개를 선택하셨지만 실제 추출된 섹션은 "
                f"**{_actual_count}개**입니다. 생성 시 {_actual_count}개 기준으로 처리됩니다."
            )

    st.markdown("---")

    # ────────────────────────────────────────
    # (B) 날짜 정보 자동 입력
    # ────────────────────────────────────────
    st.markdown("### 📅 날짜 정보")

    col_month, col_year, _ = st.columns([2, 2, 4])
    with col_month:
        # 여기를 수정하면 영문 월 기본값이 바뀝니다
        month_input = st.text_input(
            "영문 월",
            value=st.session_state.month_en,
            placeholder="예) March",
            help="표지에 표시될 영문 월 이름입니다.",
        )
        st.session_state.month_en = month_input

    with col_year:
        # 여기를 수정하면 연도 기본값이 바뀝니다
        year_input = st.text_input(
            "연도",
            value=st.session_state.year,
            placeholder="예) 2026",
            help="표지에 표시될 연도입니다.",
        )
        st.session_state.year = year_input

    st.markdown("---")

    # ────────────────────────────────────────
    # (C) 표지 이미지 선택
    # ────────────────────────────────────────
    st.markdown("### 🖼️ 표지 이미지 선택")

    images = data.get("images", [])

    if not images:
        st.warning("추출된 이미지가 없습니다. 표지 이미지 없이 생성됩니다.")
        cover_image_bytes = None
    else:
        # 가장 큰 이미지를 자동 추천 (조감도일 가능성 높음)
        largest = max(images, key=lambda x: x["width"] * x["height"])
        default_idx = largest["index"] + 1  # 1부터 시작하는 번호

        st.caption(f"총 {len(images)}개 이미지 추출됨 — 가장 큰 이미지(#{default_idx})를 자동 추천합니다.")

        # 이미지 미리보기 — 처음 8장 + 나머지는 expander
        cols = st.columns(4)
        for i, img_data in enumerate(images[:8]):
            with cols[i % 4]:
                st.image(img_data["pil_image"], use_container_width=True)
                st.caption(f"#{img_data['index']+1} ({img_data['width']}×{img_data['height']})")

        if len(images) > 8:
            with st.expander(f"나머지 {len(images)-8}개 이미지 더 보기"):
                _more_cols = st.columns(4)
                for i, img_data in enumerate(images[8:]):
                    with _more_cols[i % 4]:
                        st.image(img_data["pil_image"], use_container_width=True)
                        st.caption(f"#{img_data['index']+1} ({img_data['width']}×{img_data['height']})")

        # 이미지 번호 입력
        img_num = st.number_input(
            "표지에 사용할 이미지 번호",
            min_value=1,
            max_value=len(images),
            value=min(default_idx, len(images)),
            step=1,
            help="위 갤러리에서 원하는 이미지 번호를 입력하세요.",
        )
        st.session_state.cover_image_index = int(img_num) - 1

        # 선택된 이미지 미리보기
        selected = images[st.session_state.cover_image_index]
        st.markdown("**선택된 표지 이미지:**")
        st.image(selected["pil_image"], width=500)

        # PIL Image → bytes 변환 후 세션에 저장
        import io as _io
        buf = _io.BytesIO()
        selected["pil_image"].save(buf, format="PNG")
        cover_image_bytes = buf.getvalue()
        st.session_state.cover_image_bytes = cover_image_bytes

    st.markdown("---")

    # ────────────────────────────────────────
    # (D) 섹션 이미지 선택 (선택 사항)
    # ────────────────────────────────────────
    st.markdown("### 📌 섹션 이미지 선택 (선택 사항)")
    st.caption("섹션 구분 페이지의 원형 자리(3개)에 들어갈 이미지를 선택하세요. 4개 섹션(01~04) 모두 동일하게 적용됩니다.")

    if not images:
        st.info("추출된 이미지가 없어 섹션 이미지를 선택할 수 없습니다.")
    else:
        import io as _io
        with st.expander("섹션 이미지 선택 (원형 자리 3개)"):
            # 이미지 갤러리 — 처음 9장 + 나머지는 expander
            _gcols = st.columns(3)
            for _i, _img in enumerate(images[:9]):
                with _gcols[_i % 3]:
                    st.image(_img["pil_image"], use_container_width=True)
                    st.caption(f"#{_img['index']+1}")
            if len(images) > 9:
                with st.expander(f"나머지 {len(images)-9}개 이미지 더 보기"):
                    _more_gcols = st.columns(3)
                    for _i, _img in enumerate(images[9:]):
                        with _more_gcols[_i % 3]:
                            st.image(_img["pil_image"], use_container_width=True)
                            st.caption(f"#{_img['index']+1}")

            st.markdown("")

            # 슬롯 위치 안내 미니어처
            st.image(_make_divider_miniature(),
                     caption="원형 슬롯 위치 — ① 우측 상단 / ② 우측 하단 / ③ 중앙 좌측",
                     use_container_width=True)

            st.markdown("")

            # 3개 슬롯 번호 입력 + 미리보기 (가로 3열)
            _prev_list = st.session_state.section_img_idx_list
            _new_idx_list   = []
            _new_bytes_list = []

            _scols = st.columns(3)
            _slot_labels = ["원형 자리 1번 (우측 상단)", "원형 자리 2번 (우측 하단)", "원형 자리 3번 (중앙 좌측)"]
            for _slot in range(3):
                with _scols[_slot]:
                    st.caption(_slot_labels[_slot])
                    _prev = _prev_list[_slot] if _slot < len(_prev_list) else 0
                    _sel_num = st.number_input(
                        f"이미지 번호 (0 = 선택 안 함)",
                        min_value=0,
                        max_value=len(images),
                        value=_prev,
                        step=1,
                        key=f"sec_oval_slot_{_slot}",
                    )
                    _new_idx_list.append(int(_sel_num))
                    if _sel_num > 0:
                        _sel = images[int(_sel_num) - 1]
                        st.image(_sel["pil_image"], use_container_width=True)
                        st.caption(f"선택됨: #{int(_sel_num)}")
                        _buf = _io.BytesIO()
                        _sel["pil_image"].save(_buf, format="PNG")
                        _new_bytes_list.append(_buf.getvalue())
                    else:
                        st.caption("(선택 안 함)")
                        _new_bytes_list.append(None)

            st.session_state.section_img_idx_list   = _new_idx_list
            st.session_state.section_img_bytes_list = _new_bytes_list

    st.markdown("---")

    # ────────────────────────────────────────
    # (D-2) 목차 이미지 선택 (선택 사항)
    # ────────────────────────────────────────
    st.markdown("### 📌 목차 이미지 선택 (선택 사항)")
    st.caption("목차 페이지의 원형 자리(1개)에 들어갈 이미지를 선택하세요.")

    if not images:
        st.info("추출된 이미지가 없어 목차 이미지를 선택할 수 없습니다.")
    else:
        import io as _io
        with st.expander("목차 이미지 선택 (원형 자리 1개)"):
            _tgcols = st.columns(3)
            for _i, _img in enumerate(images[:9]):
                with _tgcols[_i % 3]:
                    st.image(_img["pil_image"], use_container_width=True)
                    st.caption(f"#{_img['index']+1}")
            if len(images) > 9:
                with st.expander(f"나머지 {len(images)-9}개 이미지 더 보기"):
                    _more_tgcols = st.columns(3)
                    for _i, _img in enumerate(images[9:]):
                        with _more_tgcols[_i % 3]:
                            st.image(_img["pil_image"], use_container_width=True)
                            st.caption(f"#{_img['index']+1}")

            st.markdown("")

            st.image(_make_toc_miniature(),
                     caption="원형 슬롯 위치 — ① 좌측 중앙",
                     use_container_width=True)

            st.markdown("")

            _toc_prev = st.session_state.toc_img_idx
            _toc_col, _ = st.columns([1, 2])
            with _toc_col:
                st.caption("원형 자리 1번 (좌측 중앙)")
                _toc_sel_num = st.number_input(
                    "이미지 번호 (0 = 선택 안 함)",
                    min_value=0,
                    max_value=len(images),
                    value=_toc_prev,
                    step=1,
                    key="toc_oval_slot_0",
                )
                st.session_state.toc_img_idx = int(_toc_sel_num)
                if _toc_sel_num > 0:
                    _toc_sel = images[int(_toc_sel_num) - 1]
                    st.image(_toc_sel["pil_image"], use_container_width=True)
                    st.caption(f"선택됨: #{int(_toc_sel_num)}")
                    _tbuf = _io.BytesIO()
                    _toc_sel["pil_image"].save(_tbuf, format="PNG")
                    st.session_state.toc_img_bytes = _tbuf.getvalue()
                else:
                    st.caption("(선택 안 함)")
                    st.session_state.toc_img_bytes = None

    st.markdown("---")

    # ────────────────────────────────────────
    # (E) 표지·목차·섹션 미리보기
    # ────────────────────────────────────────
    st.markdown("### 🚀 표지·목차·섹션 미리보기")
    st.caption("표지 1장 + 목차 1장 + 섹션 구분 페이지 4장 = 총 6장짜리 미리보기 PPT를 생성합니다. 정식 완성본이 아닙니다.")

    if st.button("🚀 표지·목차·섹션 미리 생성", type="primary"):
        try:
            with st.spinner("미리보기 PPT를 생성하는 중입니다..."):
                import io as _io
                preview_bytes = build_preview_presentation(
                    business_name=st.session_state.business_name,
                    year=st.session_state.year,
                    month_en=st.session_state.month_en,
                    cover_image_bytes=st.session_state.cover_image_bytes,
                    section_image_bytes_list=st.session_state.section_img_bytes_list,
                    toc_image_bytes_list=[st.session_state.toc_img_bytes],
                )
            _preview_name = (
                f"preview_{st.session_state.business_name}"
                f"_{datetime.now().strftime('%y%m%d')}.pptx"
            )
            st.success("✅ 미리보기 생성 완료 (표지 1장 + 목차 1장 + 섹션 구분 4장)")
            st.download_button(
                label="⬇️ 미리보기 PPT 다운로드",
                data=preview_bytes,
                file_name=_preview_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as e:
            st.error(f"❌ 미리보기 생성 중 오류가 발생했습니다: {e}")

    st.markdown("---")

    # ────────────────────────────────────────
    # (F) 하단 네비게이션
    # ────────────────────────────────────────
    col_prev, col_space, col_next = st.columns([2, 4, 2])

    with col_prev:
        if st.button("← 이전 단계", use_container_width=True):
            st.session_state.current_step = 2
            st.rerun()

    with col_next:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 4
            st.rerun()


# ─────────────────────────────────────────────
# [4단계: 전체 PPT 생성 및 다운로드]
# ─────────────────────────────────────────────
def show_step4():
    data = st.session_state.extracted_data
    if data is None:
        st.warning("추출 데이터가 없습니다. 1단계로 돌아가주세요.")
        if st.button("← 1단계로"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.markdown("## 4단계. 전체 PPT 생성")
    st.caption("지금까지 설정한 내용으로 완성된 제안서 PPT를 생성합니다.")
    st.markdown("")

    # ────────────────────────────────────────
    # (A) 파싱된 슬라이드 구성 요약
    # ────────────────────────────────────────
    pages = st.session_state.parsed_pages
    st.markdown("### 📋 슬라이드 구성 미리보기")

    if not pages:
        st.warning("⚠️ 문서에서 구조화된 내용을 찾지 못했습니다. 본문 슬라이드 없이 표지·목차·연락처만 생성됩니다.")
    else:
        st.success(f"총 **{len(pages)}개** 내용 슬라이드가 감지되었습니다.")

        # 섹션별 슬라이드 수 표시
        from collections import Counter
        section_counts = Counter(p.get("section_title", "(섹션 없음)") for p in pages)
        for sec, cnt in section_counts.items():
            st.markdown(f"- **{sec or '(섹션 없음)'}** — {cnt}개 슬라이드")

        # 상세 펼치기
        with st.expander("슬라이드 상세 내용 보기"):
            for i, page in enumerate(pages):
                st.markdown(f"**[{i+1}] {page.get('section_title', '')} / {page.get('subtitle', '')}**")
                body = page.get("body_text", "").strip()
                if body:
                    st.caption(body[:200] + ("..." if len(body) > 200 else ""))
                st.markdown("---")

    st.markdown("---")

    # ────────────────────────────────────────
    # (B) 생성 설정 요약
    # ────────────────────────────────────────
    st.markdown("### ⚙️ 생성 설정 확인")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("사업명", st.session_state.business_name or "(미입력)")
    with col2:
        st.metric("날짜", f"{st.session_state.month_en} {st.session_state.year}")
    with col3:
        cover_ok = st.session_state.cover_image_bytes is not None
        st.metric("표지 이미지", "선택됨" if cover_ok else "없음")

    st.markdown("---")

    # ────────────────────────────────────────
    # (C) PPT 생성 버튼
    # ────────────────────────────────────────
    st.markdown("### 🚀 PPT 생성")

    if st.button("📄 완성 PPT 생성하기", type="primary", use_container_width=False):
        try:
            with st.spinner("PPT를 생성하는 중입니다... 잠시만 기다려주세요."):
                ppt_bytes = build_full_presentation(
                    business_name=st.session_state.business_name,
                    year=st.session_state.year,
                    month_en=st.session_state.month_en,
                    pages=pages,
                    cover_image_bytes=st.session_state.cover_image_bytes,
                    section_image_bytes_list=st.session_state.section_img_bytes_list,
                    toc_count=st.session_state.toc_count,
                    toc_image_bytes_list=[st.session_state.toc_img_bytes],
                )

            filename = make_output_filename(st.session_state.business_name)
            st.success(f"✅ PPT 생성 완료!")

            st.download_button(
                label="⬇️ PPT 다운로드",
                data=ppt_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=False,
            )

        except Exception as e:
            st.error(f"❌ PPT 생성 중 오류가 발생했습니다: {e}")
            import traceback
            st.code(traceback.format_exc(), language="text")

    st.markdown("---")

    # ────────────────────────────────────────
    # (D) 하단 네비게이션
    # ────────────────────────────────────────
    if st.button("← 이전 단계 (설정 변경)", use_container_width=False):
        st.session_state.current_step = 3
        st.rerun()


# ─────────────────────────────────────────────
# [변환 작업 탭 전체 라우터]
# current_step 값에 따라 각 단계 화면을 표시합니다
# ─────────────────────────────────────────────
def show_conversion_tab():
    # 상단 Stepper 항상 표시
    render_stepper(st.session_state.current_step)

    if st.session_state.current_step == 1:
        show_step1()
    elif st.session_state.current_step == 2:
        show_step2()
    elif st.session_state.current_step == 3:
        show_step3()
    elif st.session_state.current_step == 4:
        show_step4()
    else:
        st.info(f"📌 {st.session_state.current_step}단계는 추후 구현 예정입니다.")


# ─────────────────────────────────────────────
# [메인 화면 함수]
# ─────────────────────────────────────────────
def show_main():
    # ── 사이드바: 로그아웃 버튼 ──
    with st.sidebar:
        st.markdown("### 레인필드투자자문 IM")
        st.markdown("---")
        # 로그아웃 버튼 (여기를 수정하면 로그아웃 버튼 텍스트가 바뀝니다)
        if st.button("🚪 로그아웃", use_container_width=True):
            # 로그아웃 시 모든 작업 데이터 초기화
            st.session_state.logged_in = False
            st.session_state.current_step = 1
            st.session_state.uploaded_file = None
            st.session_state.extracted_data = None
            st.session_state.business_name = ""
            st.session_state.toc_count = 4
            st.session_state.month_en = datetime.now().strftime("%B")
            st.session_state.year = str(datetime.now().year)
            st.session_state.cover_image_index = 0
            st.session_state.cover_image_bytes = None
            st.session_state.parsed_pages = []
            st.session_state.section_img_idx_list   = [0, 0, 0]
            st.session_state.section_img_bytes_list = [None, None, None]
            st.session_state.toc_img_idx   = 0
            st.session_state.toc_img_bytes = None
            st.rerun()  # 로그인 화면으로 전환

    # ── 메인 상단 제목 ──
    # 여기를 수정하면 메인 화면 상단 제목이 바뀝니다
    st.markdown(
        "<h1 style='text-align:center;'>🌧 레인필드투자자문 IM</h1>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── 탭 구성 ──
    # 탭 이름을 바꾸려면 아래 리스트의 문자열을 수정하세요
    tab1, tab2 = st.tabs(["🔄 변환 작업", "📁 저장된 파일 / 메모"])

    # ── 탭1: 변환 작업 ──
    with tab1:
        show_conversion_tab()

    # ── 탭2: 저장된 파일 / 메모 ──
    with tab2:
        st.info("📌 추후 STEP 7에서 구현될 예정입니다.")

    # ── 하단 푸터 ──
    # 여기를 수정하면 하단 저작권 문구가 바뀝니다
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        "<p style='text-align:center; color:lightgray; font-size:12px;'>"
        "ⓒ 2026 Rainfield Investment Advisory"
        "</p>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# [화면 분기]
# - 로그인 상태이면 메인 화면 표시
# - 로그아웃 상태이면 로그인 화면 표시
# ─────────────────────────────────────────────
if st.session_state.logged_in:
    show_main()
else:
    show_login()
