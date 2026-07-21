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
from modules.content_parser import (
    parse_document_from_bytes,
    extract_toc_map,
    extract_section_labels,
    split_into_5_sections,
    remap_pages_for_5sections,
)

# ─────────────────────────────────────────────
# [페이지 기본 설정]
# - page_title : 브라우저 탭에 표시되는 제목 (여기를 수정하면 탭 제목이 바뀝니다)
# - layout     : "wide" = 화면 전체 너비 사용 / "centered" = 가운데 정렬 좁은 화면
# - initial_sidebar_state : "collapsed" = 사이드바 기본 접힘 / "expanded" = 기본 펼침
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="IM 생성기",
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
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None          # PDF 원본 bytes(좌표기반 사진/표복원용)
if "ppt_bytes" not in st.session_state:
    st.session_state.ppt_bytes = None          # 생성된 PPT bytes(5단계 내용검수용)

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
            "<h2 style='text-align:center;'>IM 생성기</h2>",
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
    st.caption("📄 **워드(.docx)는 무조건 올려주세요 — 필수.** 본문 글자·목차 항목 추출에 꼭 필요합니다. "
               "🖼️ **PDF는 함께 올리면 더 좋아요 — 권장.** 페이지 사진을 원본 그대로 보여줍니다. "
               "(PDF / Word(.doc·.docx) 지원)")
    st.markdown("")

    # 파일 업로더 (워드·PDF 함께 업로드 가능 — 자금판과 동일 방식, 워드는 PDF로 변환)
    uploaded_files = st.file_uploader(
        label="파일을 여기에 끌어다 놓거나 클릭해서 선택하세요 (워드/PDF 함께 가능)",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True,
        key="file_uploader_widget",
    )

    if uploaded_files:
        # 워드/PDF 분리 — PDF 있으면 PDF로, 없으면 워드를 PDF로 변환해 처리
        pdf_up = next((f for f in uploaded_files if f.name.lower().endswith(".pdf")), None)
        word_up = next((f for f in uploaded_files
                        if f.name.lower().endswith((".docx", ".doc"))), None)

        # 올린 파일 정보 표시
        total_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
        if total_mb >= 50:
            st.warning("⚠️ 파일이 큽니다. 추출에 시간이 걸릴 수 있습니다.")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("올린 파일", f"{len(uploaded_files)}개")
        with col2:
            st.metric("합계 크기", f"{total_mb:.2f} MB")
        st.caption("파일: " + " · ".join(f.name for f in uploaded_files))
        if pdf_up and word_up:
            st.info("워드+PDF 둘 다 올렸습니다. 📄 **워드로 글자·항목**을 뽑고, 🖼️ **PDF로 페이지 사진**을 보여줍니다.")
        elif word_up and not pdf_up:
            st.info("워드만 올렸습니다. 자동으로 PDF로 변환해 사진까지 보여줍니다. "
                    "(원본 PDF도 함께 올리면 사진이 더 원본에 가까워요)")
        elif pdf_up and not word_up:
            st.warning("⚠️ **워드(.docx)가 없습니다 — 워드는 필수예요.** "
                       "PDF만으로는 글자·목차 항목이 안 잡힐 수 있습니다(내용이 이미지로 된 PDF가 많음). "
                       "워드를 꼭 함께 올려주세요.")

        st.markdown("")

        # 추출 시작 버튼
        if st.button("🔍 추출 시작", type="primary", use_container_width=False):
            try:
                # 글자·목차 항목은 '워드(변환 PDF)'에서 정확히 뽑고, 사진은 '원본 PDF'로 보여준다.
                proc_pdf = None      # 글자/항목 추출용
                conv_err = None
                proc_name = word_up.name if word_up is not None else (pdf_up.name if pdf_up else "")
                if word_up is not None:
                    from modules.preview import convert_word_to_pdf
                    with st.spinner("워드를 PDF로 변환하는 중입니다... (자금판과 동일 방식)"):
                        proc_pdf, conv_err = convert_word_to_pdf(word_up.getvalue(), word_up.name)
                    if not proc_pdf and pdf_up is not None:      # 변환 실패 시 PDF로
                        proc_pdf = pdf_up.getvalue()
                        proc_name = pdf_up.name
                elif pdf_up is not None:
                    proc_pdf = pdf_up.getvalue()

                with st.spinner("파일에서 텍스트와 이미지를 추출하는 중입니다..."):
                    if proc_pdf is not None:
                        result = extract_from_pdf(proc_pdf)
                        st.session_state.pdf_bytes = proc_pdf        # 좌표기반 표·이미지 복원용
                        parse_bytes, parse_name = proc_pdf, "converted.pdf"
                        # 사진 표시용 PDF: 원본 PDF 우선(글자 PDF와 페이지 수 같을 때만 정렬 유지)
                        view_pdf = proc_pdf
                        if pdf_up is not None:
                            _up = pdf_up.getvalue()
                            if _pdf_page_count(_up) == _pdf_page_count(proc_pdf):
                                view_pdf = _up
                        st.session_state.view_pdf_bytes = view_pdf
                    else:
                        # 변환 실패 폴백: python-docx 텍스트만(레이아웃·사진 제한)
                        result = extract_from_docx(word_up.getvalue())
                        st.session_state.pdf_bytes = None
                        st.session_state.view_pdf_bytes = None
                        parse_bytes, parse_name = word_up.getvalue(), word_up.name
                        if conv_err:
                            st.warning(f"워드→PDF 변환 실패 — 텍스트만 추출(사진 제한). 원인: {conv_err}")

                # 추출 결과 저장
                st.session_state.extracted_data = result
                st.session_state.uploaded_file  = proc_name

                # 사업명 자동 감지
                detected = detect_business_name(result["pages_text"])
                st.session_state.business_name = detected

                # 슬라이드 구조 파싱 (content_parser) — 변환된 PDF 기준
                try:
                    st.session_state.parsed_pages = parse_document_from_bytes(
                        parse_bytes, parse_name
                    )
                except Exception:
                    st.session_state.parsed_pages = []

                # 새 문서이므로 이전 목차 편집·이미지 상태 초기화(3단계 새로 시작)
                for _k in ("toc_edit", "toc_hashtags", "toc_page_cache", "toc_view_page",
                           "view_pdf_bytes"):
                    st.session_state.pop(_k, None)
                _clear_toc_widget_state()

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

    st.markdown("## 4단계. 레이아웃 및 표지 미리 생성")
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
        index=0,                       # 기본 선택값 = '목차 4개 형식'(자동추천 문구는 위에 그대로 유지)
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

        # ★섹션/목차 이미지 선택과 동일하게 expander(접기/펼치기)로 통일
        with st.expander("표지 이미지 선택"):
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
            st.session_state.current_step = 3
            st.rerun()

    with col_next:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 5
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

    st.markdown("## 5단계. 전체 PPT 생성")
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
                # ── 4/5형식 분기: 5형식이면 자동 분할 후 pages·toc_map 재구성 ──
                toc_choice = st.session_state.toc_count
                if pages and toc_choice == 5:
                    try:
                        _toc4   = extract_toc_map(pages)
                        _lbl4   = extract_section_labels(pages)
                        _toc5, _lbl5, _split = split_into_5_sections(_toc4, _lbl4)
                        _pages5 = remap_pages_for_5sections(pages, _split, _lbl5)
                        final_pages     = _pages5
                        final_toc_map   = dict(_toc5)
                        final_toc_map["_labels"] = _lbl5  # 섹션 제목 레이블 주입
                        final_toc_count = 5
                    except Exception as _e:
                        st.warning(f"⚠️ 5섹션 자동 분할 실패({_e}). 4섹션으로 생성합니다.")
                        final_pages     = pages
                        final_toc_map   = None
                        final_toc_count = 4
                else:
                    final_pages     = pages
                    final_toc_map   = None   # build_full_presentation 내부에서 자동 추출
                    final_toc_count = toc_choice

                # ── ★LLM 구조화 파이프라인 적용(테스트 하니스와 동일) ─────────────
                #   본문 슬라이드를 LLM 페이지 구조화 경로로 생성하고, Executive
                #   Summary 는 '지금 업로드한 PDF' 의 ES 페이지만 보고 생성한다.
                os.environ["RAINFIELD_LLM"] = "1"   # _build_content_block LLM 경로 ON
                exec_summary_data = None
                try:
                    from modules.llm_structure import enrich_and_number
                    # ★PDF 원본을 임시파일로 넘겨 좌표기반 사진(조감도)·표복원(매매사례)·썸네일 활성화
                    _pdf_b = st.session_state.get("pdf_bytes")
                    _pdf_path = None
                    if _pdf_b:
                        import tempfile
                        _tf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                        _tf.write(_pdf_b); _tf.close(); _pdf_path = _tf.name
                    enrich_and_number(final_pages, pdf_path=_pdf_path)
                    from modules.ai_slide_builders import generate_executive_summary
                    _es_text = "\n".join(
                        (p.get("raw_text", "") or "") for p in final_pages
                        if "Executive Summary" in (p.get("raw_text", "") or ""))
                    if not _es_text:    # 폴백: 앞 2페이지
                        _es_text = "\n".join((p.get("raw_text", "") or "")
                                             for p in final_pages[:2])
                    _es = generate_executive_summary(_es_text)
                    if _es.get("ok"):
                        exec_summary_data = _es["data"]
                except Exception as _llm_e:
                    st.warning(f"⚠️ LLM 구조화 일부 실패({_llm_e}) — 기본 경로로 진행합니다.")

                ppt_bytes = build_full_presentation(
                    business_name=st.session_state.business_name,
                    year=st.session_state.year,
                    month_en=st.session_state.month_en,
                    pages=final_pages,
                    cover_image_bytes=st.session_state.cover_image_bytes,
                    section_image_bytes_list=st.session_state.section_img_bytes_list,
                    toc_count=final_toc_count,
                    toc_map=final_toc_map,
                    toc_image_bytes_list=[st.session_state.toc_img_bytes],
                    exec_summary_data=exec_summary_data,
                )

            # ★5단계 내용검수에서 쓰도록 생성 PPT 보관
            st.session_state.ppt_bytes = ppt_bytes
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
        st.session_state.current_step = 4
        st.rerun()


# ─────────────────────────────────────────────
# [3단계: 목차 구성] — 자동 추출 목차 편집 + 원본 항목(해시태그) 매치. (읽기/편집만, 이미지·재배치 X)
# ─────────────────────────────────────────────
def _clean_toc_text(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "").strip())


def _is_toc_heading(t):
    """목차 제목/소제목다운 '짧은 문구'인지 — 긴 본문 문장/서술형은 제외."""
    t = (t or "").strip()
    if not (0 < len(t) <= 45):
        return False
    if re.search(r"(습니다|입니다|합니다|이다|였다|된다|한다|하였다)\.?$", t):
        return False
    return True


# 항목 머리글 패턴: (1) / 1) / 1. / 1.1 / ①~⑳ / ■▪●▶ / (가) / 가. / I. II.
_ITEM_MARK = re.compile(
    r"^\s*(\(?\d{1,2}[\).]|\d{1,2}\.\d{1,2}|[①-⑳]|[■▪●◦▶]|\([가-힣]\)|[가-힣]\.|[IVXivx]{1,4}\.)")


def _page_items(pages_text, page):
    """지정 페이지의 텍스트에서 '항목 머리글'로 시작하는 짧은 줄만 뽑음(그 페이지의 (1)(2)(3)…)."""
    if not pages_text or not (1 <= page <= len(pages_text)):
        return []
    out, seen = [], set()
    for ln in (pages_text[page - 1] or "").split("\n"):
        s = re.sub(r"\s+", " ", ln.strip())
        if 2 < len(s) <= 60 and _ITEM_MARK.match(s) and s not in seen:
            seen.add(s)
            out.append(s)
    return out


def _build_toc_from_parsed(parsed_pages):
    """parsed_pages로 목차 자동 추출 시도(best-effort) → groups.
       ※IM마다 섹션/소제목 형식이 제각각이고 러닝헤더(사업명)가 섞여 신뢰도 낮음 → '시도'용.
       여러 페이지에 반복되는 제목(러닝헤더)은 빈도로 걸러낸다."""
    from collections import OrderedDict, Counter
    freq = Counter(_clean_toc_text(pd.get("section_title") or pd.get("section_label"))
                   for pd in (parsed_pages or []))
    npages = max(1, len(parsed_pages or []))

    def _bad(t):
        return (not _is_toc_heading(t)) or freq.get(t, 0) > max(3, npages * 0.4)

    sections = OrderedDict()
    for pd in (parsed_pages or []):
        sec = _clean_toc_text(pd.get("section_title") or pd.get("section_label"))
        sub = _clean_toc_text(pd.get("subtitle"))
        if _bad(sec):
            sec = ""
        sub = sub if _is_toc_heading(sub) else ""
        if sec:
            sections.setdefault(sec, [])
            if sub and sub != sec and sub not in sections[sec]:
                sections[sec].append(sub)
    return [{"title": k, "subs": [{"text": s, "page": 0, "item": ""} for s in v]}
            for k, v in sections.items()]


def _adjust_groups(groups, n):
    """대분류 개수를 정확히 n개로(부족하면 빈 그룹 추가, 많으면 뒤에서 제거)."""
    while len(groups) < n:
        groups.append({"title": "", "subs": [{"text": "", "page": 0, "item": ""}]})
    while len(groups) > n:
        groups.pop()
    return groups


def _clear_toc_widget_state():
    """목차 편집 위젯 상태 초기화(구조 변경 후 재시딩용)."""
    for k in list(st.session_state.keys()):
        if k.startswith(("toctitle_", "tocsub_", "tocpage_", "tocitem_")):
            del st.session_state[k]


def _init_toc_edit(parsed, auto=False):
    """목차 편집 상태 초기화 — 기본은 빈 4형식(직접 작성). auto=True면 자동 추출 시도."""
    groups = _build_toc_from_parsed(parsed) if auto else []
    fmt = max(4, min(5, len(groups) or 4))
    _adjust_groups(groups, fmt)
    st.session_state.toc_edit = {"format": fmt, "groups": groups}
    st.session_state.toc_count = fmt


def _pdf_page_count(pdf_bytes):
    """PDF 페이지 수(실패 시 -1) — 사진 PDF와 글자 PDF 정렬 확인용."""
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            return doc.page_count
        finally:
            doc.close()
    except Exception:
        return -1


def _render_pdf_page(pdf_bytes, page_num, zoom=2.0):
    """PDF 한 페이지를 PNG bytes로 렌더 — PyMuPDF(fitz), 자금판(전) 방식.
       외부 프로그램(LibreOffice/poppler) 없이 빠르게 렌더. 반환: (png|None, err|None)."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            pix = doc[page_num - 1].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            return pix.tobytes("png"), None
        finally:
            doc.close()
    except Exception as e:
        return None, f"이미지 렌더 실패: {e}"


def show_step_toc():
    st.markdown("## 3단계. 목차 구성")
    st.caption("왼쪽에서 원본 IM을 표지부터 넘겨보고, 오른쪽에서 목차 제목·소제목을 직접 만드세요. "
               "소제목마다 **원본 IM의 몇 페이지인지** 번호로 지정하면 됩니다. "
               "(IM마다 목차/소제목 형식이 제각각이라 자동추출은 부정확 → 페이지 번호로 매칭)")

    parsed = st.session_state.get("parsed_pages") or []
    pages_text = (st.session_state.get("extracted_data") or {}).get("pages_text", [])
    total = len(pages_text)
    pdf_bytes = st.session_state.get("pdf_bytes")           # 글자/항목 소스(워드 변환)
    view_pdf = st.session_state.get("view_pdf_bytes") or pdf_bytes   # 사진(원본 PDF 우선)
    if "toc_edit" not in st.session_state:
        _init_toc_edit(parsed)
    toc = st.session_state.toc_edit

    bc1, bc2, _bc = st.columns([1, 1, 3])
    with bc1:
        if st.button("🔄 자동 추출 시도", use_container_width=True,
                     help="원본에서 목차를 자동 추출해봅니다(IM에 따라 부정확할 수 있음)."):
            _init_toc_edit(parsed, auto=True)
            _clear_toc_widget_state()
            st.rerun()
    with bc2:
        if st.button("🧹 모두 비우기", use_container_width=True, help="빈 목차로 초기화합니다."):
            _init_toc_edit(parsed, auto=False)
            _clear_toc_widget_state()
            st.rerun()

    img_col, form_col = st.columns([1, 1])

    # ── 왼쪽: 원본 IM 전체 페이지 뷰어 (표지부터 한 장씩, 버튼 없이 바로 표시) ──
    with img_col:
        st.markdown("#### 📄 원본 IM (표지부터 한 장씩)")
        if not pdf_bytes or total == 0:
            st.warning("원본 PDF가 없어 이미지를 못 띄웁니다. 1단계에서 PDF(또는 워드→자동 PDF변환)로 올리세요.")
        else:
            cur = min(max(1, st.session_state.get("toc_view_page", 1)), total)
            pc1, pc2, pc3 = st.columns([1, 2, 1])
            with pc1:
                if st.button("◀ 이전", use_container_width=True, disabled=(cur <= 1)):
                    st.session_state.toc_view_page = cur - 1
                    st.rerun()
            with pc2:
                st.markdown(f"<div style='text-align:center; padding-top:6px;'>"
                            f"<b>{cur}</b> / {total} 페이지</div>", unsafe_allow_html=True)
            with pc3:
                if st.button("다음 ▶", use_container_width=True, disabled=(cur >= total)):
                    st.session_state.toc_view_page = cur + 1
                    st.rerun()
            # 현재 페이지를 fitz로 즉시 렌더(캐시) — 생성 버튼·다운로드 없이 바로 표시
            cache = st.session_state.setdefault("toc_page_cache", {})
            if cur not in cache:
                cache[cur] = _render_pdf_page(view_pdf, cur)
            png, err = cache[cur]
            if err:
                st.error(err)
            elif png:
                st.image(png, use_container_width=True)

    # ── 오른쪽: 목차 편집 폼 (컬럼 2중첩 방지 → 세로로 쌓음) ──
    with form_col:
        st.markdown("#### 📝 목차 구성 편집")
        fmt_sel = st.radio("목차 형식 (대분류 개수)", [4, 5],
                           index=(0 if toc["format"] == 4 else 1),
                           horizontal=True, format_func=lambda x: f"{x}형식",
                           key="toc_fmt_radio")
        if fmt_sel != toc["format"]:
            toc["format"] = fmt_sel
            _adjust_groups(toc["groups"], fmt_sel)
            st.session_state.toc_count = fmt_sel
            _clear_toc_widget_state()
            st.rerun()

        _cur = min(max(1, st.session_state.get("toc_view_page", 1)), total or 1)
        st.caption(f"소제목마다 '원본 페이지'를 넣으세요 (0=미지정). 지금 왼쪽에 보이는 페이지: **{_cur}p**")

        pending = {"del": None, "add": None}
        for gi, g in enumerate(toc["groups"]):
            with st.container(border=True):
                tk = f"toctitle_{gi}"
                st.session_state.setdefault(tk, g.get("title", ""))
                g["title"] = st.text_input(f"목차 {gi + 1} 제목", key=tk,
                                           placeholder="예: 4 리스크분석 / Executive Summary")
                for si, sub in enumerate(g["subs"]):
                    sk = f"tocsub_{gi}_{si}"
                    st.session_state.setdefault(sk, sub.get("text", ""))
                    sub["text"] = st.text_input(f"└ 소제목 {si + 1}", key=sk,
                                                placeholder="예: (1) 낙찰사례 및 환가 분석")
                    pk = f"tocpage_{gi}_{si}"
                    st.session_state.setdefault(pk, int(sub.get("page", 0) or 0))
                    sub["page"] = st.number_input(
                        "└ 원본 페이지 (0=미지정)", min_value=0,
                        max_value=(total if total else 999), step=1, key=pk)
                    # 그 페이지의 항목((1)(2)(3)…)이 잡히면 골라서 '어느 부분'인지 지정
                    items = _page_items(pages_text, int(sub.get("page") or 0))
                    if items:
                        ik = f"tocitem_{gi}_{si}"
                        opts = ["(페이지 전체)"] + items
                        if sub.get("item") and sub["item"] in opts and ik not in st.session_state:
                            st.session_state[ik] = sub["item"]
                        if ik in st.session_state and st.session_state[ik] not in opts:
                            del st.session_state[ik]
                        picked = st.selectbox("└ 그 페이지의 어느 항목?", opts, key=ik,
                                              help="이 소제목이 그 페이지의 어느 항목인지. (1)만 고르면 (1)만 가져옵니다.")
                        sub["item"] = "" if picked == "(페이지 전체)" else picked
                    else:
                        sub["item"] = sub.get("item", "")
                    if st.button("🗑 소제목 삭제", key=f"tocdel_{gi}_{si}"):
                        pending["del"] = (gi, si)
                if st.button("➕ 소제목 추가", key=f"tocadd_{gi}"):
                    pending["add"] = gi

        if pending["del"] is not None:
            gi, si = pending["del"]
            toc["groups"][gi]["subs"].pop(si)
            _clear_toc_widget_state()
            st.rerun()
        if pending["add"] is not None:
            toc["groups"][pending["add"]]["subs"].append({"text": "", "page": 0, "item": ""})
            _clear_toc_widget_state()
            st.rerun()

    # ── 저장 요약 + 네비게이션 (전체 폭) ──
    st.markdown("---")
    total_subs = sum(len(g["subs"]) for g in toc["groups"])
    with_page = sum(1 for g in toc["groups"] for s in g["subs"] if s.get("page"))
    st.success(f"저장됨 · 대분류 {toc['format']}개 · 소제목 {total_subs}개 · 페이지 지정 {with_page}건")
    nc1, _ns, nc2 = st.columns([2, 4, 2])
    with nc1:
        if st.button("← 이전 단계", use_container_width=True):
            st.session_state.current_step = 2
            st.rerun()
    with nc2:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 4
            st.rerun()


def _find_toc_pages(pages_text):
    """원본 텍스트에서 '목차' 페이지(1-based) 추정 — 상단에 '목차/CONTENTS' 표기가 있는 페이지."""
    hits = []
    for i, t in enumerate(pages_text or [], start=1):
        head = (t or "")[:300]
        low = head.lower()
        if "목차" in head or "contents" in low or "table of contents" in low:
            hits.append(i)
    return hits


# ─────────────────────────────────────────────
# ['처음으로' — 현재 작업 전체 초기화 후 1단계로]
# ─────────────────────────────────────────────
def _reset_to_start():
    """로그인 상태만 남기고 모든 세션 상태(업로드·추출·선택·생성 PPT 등)를 비운 뒤 1단계로.
       삭제된 키들은 스크립트 상단의 기본값 초기화 블록에서 다음 rerun 때 재생성된다.
       (메모는 memos.json에 저장돼 있어 세션 초기화와 무관하게 그대로 유지됨)"""
    for k in list(st.session_state.keys()):
        if k != "logged_in":
            del st.session_state[k]
    st.session_state.current_step = 1
    st.rerun()


def _confirm_home_dialog():
    """st.dialog 지원 시 확인 팝업으로 '처음으로' 확인/취소."""
    @st.dialog("처음으로")
    def _dlg():
        st.write("정말 처음으로 가시겠어요? 현재 작업이 초기화됩니다.")
        c1, c2 = st.columns(2)
        if c1.button("확인", type="primary", use_container_width=True):
            _reset_to_start()
        if c2.button("취소", use_container_width=True):
            st.rerun()        # 팝업만 닫고 현재 상태 유지
    _dlg()


# ─────────────────────────────────────────────
# [변환 작업 탭 전체 라우터]
# current_step 값에 따라 각 단계 화면을 표시합니다
# ─────────────────────────────────────────────
def show_conversion_tab():
    # 상단: '처음으로' 버튼 + Stepper
    _has_dialog = hasattr(st, "dialog")
    _hc1, _hc2 = st.columns([5, 1])
    with _hc2:
        if st.button("🏠 처음으로", use_container_width=True, help="현재 작업을 초기화하고 1단계로 돌아갑니다."):
            if _has_dialog:
                _confirm_home_dialog()
            else:
                st.session_state._show_home_confirm = True

    # st.dialog 미지원 환경: 인라인 확인 영역
    if not _has_dialog and st.session_state.get("_show_home_confirm"):
        st.warning("정말 처음으로 가시겠어요? 현재 작업이 초기화됩니다.")
        _cc1, _cc2 = st.columns(2)
        if _cc1.button("확인", type="primary", use_container_width=True):
            _reset_to_start()
        if _cc2.button("취소", use_container_width=True):
            st.session_state._show_home_confirm = False
            st.rerun()

    # 상단 Stepper 항상 표시
    render_stepper(st.session_state.current_step)

    if st.session_state.current_step == 1:
        show_step1()
    elif st.session_state.current_step == 2:
        show_step2()
    elif st.session_state.current_step == 3:
        show_step_toc()
    elif st.session_state.current_step == 4:
        show_step3()
    elif st.session_state.current_step == 5:
        show_step4()
    else:
        st.info(f"📌 {st.session_state.current_step}단계는 추후 구현 예정입니다.")


# ─────────────────────────────────────────────
# [메인 화면 함수]
# ─────────────────────────────────────────────
def show_main():
    # ── 사이드바: 로그아웃 버튼 ──
    with st.sidebar:
        st.markdown("### IM 생성기")
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
        "<h1 style='text-align:center;'>IM 생성기</h1>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── 변환 작업 화면 ──
    show_conversion_tab()

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
# - 비밀번호 제거: 접근 코드 없이 항상 메인 화면 표시
#   (통합 대시보드에서만 비밀번호로 접근을 통제합니다)
# - 다시 로그인 화면을 켜려면 아래 두 줄을 지우고 원래 분기(if ...: show_main() else: show_login())로 되돌리면 됩니다.
# ─────────────────────────────────────────────
show_main()
