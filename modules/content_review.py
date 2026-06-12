"""5단계 '내용 검수' — 생성된 PPT 본문이 원본 PDF/워드 내용과 일치하는지 자동 점검.

★읽기 전용: 본문 생성 로직을 건드리지 않고, 문제를 '찾아서 보여주기만' 한다(자동 수정 X).
점검 항목
  1) 내용 1:1 대조 — 원본의 '중요 숫자/금액'이 PPT에 빠졌는지(누락) + 표 빈 셀
  2) 오타/맞춤법 — PPT 한글 텍스트의 맞춤법 오류(hanspell 사용 가능 시, 아니면 규칙 기반 최소 점검)
반환: [{"page": "원본 3p"|"슬라이드 7", "type": "내용 누락"|"빈 표 셀"|"맞춤법",
        "content": "...", "suggestion": "..."}] 리스트.
"""
import io
import re

from pptx import Presentation

# 콤마/소수 포함 숫자
_NUM_RE = re.compile(r"\d[\d,\.]*\d|\d")
# 금액/비율/면적 등 '단위가 붙은 중요 값'
_UNIT_RE = re.compile(r"(억원|억|만원|만|천원|원|％|%|평|㎡|세대|개월|개동|동|호|건|명|위)")


def _norm_num(s: str) -> str:
    """비교용 숫자 정규화 — 콤마·공백 제거, 끝자리 0/소수점 정리."""
    return s.replace(",", "").replace(" ", "").rstrip("0").rstrip(".")


# ──────────────────────────────────────────────────────────
# PPT 텍스트 추출
# ──────────────────────────────────────────────────────────
def _ppt_full_text(prs) -> str:
    parts = []
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text:
                parts.append(sh.text_frame.text)
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        if c.text:
                            parts.append(c.text)
    return "\n".join(parts)


def _ppt_korean_lines(prs):
    """(슬라이드번호, 텍스트) — 맞춤법 검사용 한글 포함 줄."""
    out = []
    for i, s in enumerate(prs.slides, start=1):
        for sh in s.shapes:
            if sh.has_text_frame:
                for ln in sh.text_frame.text.split("\n"):
                    t = ln.strip()
                    if len(t) >= 6 and re.search(r"[가-힣]", t):
                        out.append((i, t))
    return out


# ──────────────────────────────────────────────────────────
# 1) 내용 1:1 대조 — 원본의 중요 숫자가 PPT에 없으면 '누락'
# ──────────────────────────────────────────────────────────
def _significant_numbers(text: str):
    """원본에서 '빠지면 안 되는 값'만 추출 — 콤마 큰수(1,290)·단위 붙은 값(980억/61%/35평).
       날짜(2026.03)·단순 정수·페이지번호 등은 표기변형/비1:1이라 노이즈 → 제외."""
    found = []
    for m in _NUM_RE.finditer(text):
        tok = m.group(0)
        digits = tok.replace(",", "").replace(".", "")
        if len(digits) < 2:
            continue
        tail = text[m.end():m.end() + 3].lstrip()
        unit_m = _UNIT_RE.match(tail)
        if not (("," in tok) or unit_m):                  # 콤마 큰수 or 단위 붙은 값만
            continue
        disp = tok + (unit_m.group(1) if unit_m else "")
        found.append((tok, disp))
    return found


def find_missing_content(pages_text, prs):
    """원본 페이지별 중요 숫자 중 PPT에 없는 것 → 누락 리스트."""
    ppt_nums = {_norm_num(m) for m in _NUM_RE.findall(_ppt_full_text(prs))}
    rows, seen = [], set()
    for pi, ptext in enumerate(pages_text or [], start=1):
        page_seen = set()
        for tok, disp in _significant_numbers(ptext):
            core = _norm_num(tok)
            if not core or core in page_seen:
                continue
            page_seen.add(core)
            if core in ppt_nums:
                continue
            key = (pi, core)
            if key in seen:
                continue
            seen.add(key)
            rows.append({"page": f"원본 {pi}p", "type": "내용 누락",
                         "content": disp, "suggestion": "PPT에 해당 수치가 없음 — 누락 여부 확인"})
    return rows


# ──────────────────────────────────────────────────────────
# 2) 빈 표 셀 — 병합(연속)칸 제외, 헤더행 제외한 진짜 빈 칸
# ──────────────────────────────────────────────────────────
def _is_merge_continuation(cell):
    tc = cell._tc
    return (tc.get("hMerge") in ("1", "true") or tc.get("vMerge") in ("1", "true"))


def find_empty_cells(prs):
    """병합 연속칸·헤더·설계상 빈 열은 제외하고, '대부분 채워진 열에 뚫린 빈칸'(데이터 누락 의심)만 표시.
       — 발코니확장/근린생활시설처럼 원래 비는 칸(열 자체가 듬성)은 노이즈라 걸러냄."""
    rows = []
    for i, s in enumerate(prs.slides, start=1):
        for sh in s.shapes:
            if not sh.has_table:
                continue
            t = sh.table
            ncol, nrow = len(t.columns), len(t.rows)
            ndata = max(1, nrow - 1)
            # 열별 채움률(데이터행 기준)
            filled = [0] * ncol
            for ci in range(ncol):
                for ri in range(1, nrow):
                    if t.rows[ri].cells[ci].text.strip():
                        filled[ci] += 1
            for ri in range(1, nrow):
                _c0 = t.rows[ri].cells[0].text.strip()
                if any(k in _c0 for k in ("합계", "소계", "합 계", "소 계", "총계", "총합계", "총 계")):
                    continue                              # 합계/소계행은 빈칸이 정상 → 제외
                # 행 자체가 듬성(데이터칸 절반 이상 빔)하면 비용항목 등 설계상 빈행 → 제외
                _rowfill = sum(1 for cj in range(1, ncol) if t.rows[ri].cells[cj].text.strip())
                if _rowfill < (ncol - 1) * 0.5:
                    continue
                for ci, c in enumerate(t.rows[ri].cells):
                    if c.text.strip() or _is_merge_continuation(c):
                        continue
                    if filled[ci] < ndata * 0.6:          # 열 자체가 듬성하면(설계상) 제외
                        continue
                    lbl = t.rows[ri].cells[0].text.strip() or (
                        t.rows[0].cells[ci].text.strip() if ncol > ci else "") or "?"
                    rows.append({"page": f"슬라이드 {i}", "type": "빈 표 셀",
                                 "content": f"'{lbl[:18]}' 행({ri+1}행 {ci+1}열) 비어있음",
                                 "suggestion": "원본에 값이 있는지 확인"})
    return rows


# ──────────────────────────────────────────────────────────
# 3) 오타 / 맞춤법
# ──────────────────────────────────────────────────────────
def check_spelling(prs):
    """맞춤법 검사 — hanspell(네이버) 설치/동작 시에만 수행. 없으면 '건너뛰기'.
       ★py-hanspell은 클라우드 설치가 자주 실패하므로 requirements에 넣지 않는다.
         라이브러리 미설치/호출 실패 시 빈 결과 + '비활성' 표시 → 앱은 절대 죽지 않음.
         (내용 1:1 대조·빈 표 셀 검출은 이 함수와 무관하게 정상 동작)"""
    try:
        from hanspell import spell_checker            # 설치돼 있을 때만(기본 미설치)
    except Exception:
        return [], "맞춤법 검사 비활성(라이브러리 미설치)"

    rows, ok = [], False
    for i, t in _ppt_korean_lines(prs):
        try:
            res = spell_checker.check(t[:480])         # hanspell 길이 제한 대비
        except Exception:
            continue
        ok = True
        errs = getattr(res, "errors", 0)
        checked = getattr(res, "checked", "")
        if errs and checked and checked.strip() != t.strip():
            rows.append({"page": f"슬라이드 {i}", "type": "맞춤법",
                         "content": t[:40], "suggestion": checked[:60]})
    if not ok:
        return [], "맞춤법 검사 비활성(호출 실패 — 네트워크)"
    return rows, "hanspell(네이버 맞춤법)"


# ──────────────────────────────────────────────────────────
# 통합 검수
# ──────────────────────────────────────────────────────────
def review_presentation(ppt_bytes: bytes, pages_text):
    """PPT bytes + 원본 페이지텍스트 → 검수 결과 dict."""
    prs = Presentation(io.BytesIO(ppt_bytes))
    missing = find_missing_content(pages_text, prs)
    empty = find_empty_cells(prs)
    spelling, spell_engine = check_spelling(prs)
    items = missing + empty + spelling
    return {
        "ok": len(items) == 0,
        "items": items,
        "counts": {"내용 누락": len(missing), "빈 표 셀": len(empty), "맞춤법": len(spelling)},
        "spell_engine": spell_engine,
    }
