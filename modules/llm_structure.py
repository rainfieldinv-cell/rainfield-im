"""
llm_structure.py
─────────────────────────────────────────────────────────
PDF 한 페이지의 원문 텍스트를 LLM(Claude)이 읽어, 가로 A4 제안서 슬라이드 1장의
구조화된 내용(JSON)으로 재구성한다. 규칙기반 추출이 한국어 IM에서 제목·본문·표를
잘못 뽑는 문제(과분할/제목오인식)를 근본 해결하기 위한 경로.

사람이 만든 대전 제안서([[rainfield-daejeon-skeleton]]) 구성 방식을 따른다:
  섹션라벨 / 소제목 / 인트로(별도 글상자) / 표(라벨-값 또는 그리드) / 불릿 / 출처

claude_api.call_claude()를 그대로 재사용(캐시·재시도·비용로깅·숫자환각검증 내장).
"""

from modules.claude_api import call_claude, verify_numbers_in_pdf

PROMPT_VERSION = "structure_v4"

# 레인필드 표준 4섹션 (고정). 섹션라벨은 항상 이 이름으로 표기.
SECTION_NAMES = {1: "사모사채 개요", 2: "금융개요", 3: "본건 사업 개요", 4: "Appendix"}

SYSTEM_PROMPT = """당신은 부동산 금융 IM(PDF)을 가로 A4 제안서 PPT로 재구성하는 전문 애널리스트입니다.
주어진 PDF '한 페이지'의 원문 텍스트를 읽고, 슬라이드 1장의 구조화된 내용을 순수 JSON으로만 출력하세요.

[제안서는 항상 4개 섹션으로 구성됩니다]
  1 = 사모사채 개요   : Executive Summary, 본건 거래/딜 요약, 핵심 투자포인트
  2 = 금융개요        : 금융조건/대출조건(트랜치·금리·LTV·수수료), 투자구조도, 기초자산, 기한이익상실 등 약정조건
  3 = 본건 사업 개요  : 담보/토지 개요, 사업개요(건축·분양), 사업일정·에쿼티, 사업수지, 입지·시장·분양사례 분석, 차주/시행사 개요·재무
  4 = Appendix        : 현장사진, 시공사 개요, 관계사 개요, 사업계획승인서 등 인허가 첨부, 기타 부록
이 페이지가 어느 섹션에 속하는지 1~4 중 하나로 판단하세요.

[출력 JSON 스키마]
{
  "section_num": 1,
  "subtitle": "이 페이지의 소제목(번호 없이 이름만). 예 '담보토지 개요', '건축개요', '주요 금융조건'.",
  "intro": "본문 도입 1~2문장. 페이지번호·머리말·footer·출처는 절대 포함하지 말 것. 없으면 \\"\\".",
  "bullets": ["표로 만들 수 없는 본문/설명 항목들. 각 항목 한 문장. 없으면 빈 배열."],
  "tables": [
    {
      "title": "표 제목(표 위 작은 라벨). 없으면 \\"\\".",
      "kind": "label_value | grid",
      "header": ["grid일 때만 열 제목 배열. label_value면 빈 배열."],
      "rows": [["..."], ["..."]]
    }
  ],
  "source": "출처 문구. 예 '국토부 실거래가, KB부동산'. 없으면 \\"\\"."
}

[가장 중요 — 표는 "요약"이 아니라 "원문 그대로 재현"]
★원문 페이지의 표는 빠짐없이 그대로 재현하세요. 요약하거나, 일부(형광펜·강조 부분)만 표로 만들지 마세요.
  원문에 표가 있으면 그 표의 모든 행·모든 열을 tables에 그대로 옮깁니다.
★본문을 bullets로 "요약"하지 마세요. 요약은 오직 intro(1~2문장)에만. 표로 표현되는 내용은 전부 tables로.
  bullets는 표도 아니고 도입문도 아닌 진짜 설명문/주석만(거의 비어 있어야 정상).

[중요 규칙]
1. section_num은 반드시 1~4 정수. subtitle에는 절대 번호(1.1 등)를 붙이지 말고 이름만(번호는 시스템이 부여).
2. label_value: 좌측 항목명(구분) + 우측 값(내용)인 2열 표. rows=[["대지면적","21,665평"], ...]. header는 빈 배열.
   ★★한 구분(라벨)에 여러 세부항목(•, ①②③, 1)2)3), - 등)이 딸리면 **절대 항목마다 행을 나누지 말고**,
     그 구분을 한 행으로 두고 **모든 세부항목을 내용 칸 하나에 줄바꿈(\\n)으로 이어서** 넣으세요.
     세부항목 번호(①②③)를 구분 칸에 넣지 마세요 — 구분 칸에는 반드시 원문의 항목명(예 '주요 채권보전조치',
     '인출선행조건', '기한이익상실 사유', '대주의 의사결정')을 그대로 쓰고, 내용 칸에 그 아래 모든 줄을 넣습니다.
     예) 원문:  주요 채권보전조치 / • A / • B / • C
        → rows에 ["주요 채권보전조치", "• A\\n• B\\n• C"]  (행 1개. A·B·C로 행 3개로 쪼개지 말 것)
     '1) 당연…', '2) 선택적…' 같은 소제목도 내용 칸 안에 줄바꿈으로 함께 넣으세요(구분 칸으로 빼지 말 것).
3. grid: 다열 표(트랜치·Cash-In/Out·재무표처럼 행·열이 격자인 표). header에 열 제목, rows에 데이터행.
   원문 표의 열 구성을 그대로 따르세요. ★단, '구분|내용' 처럼 좌측이 항목명·우측이 설명인 2열은 grid가 아니라 label_value.
4. ★병합/반복: 원문에서 세로로 같은 값이 연속 반복되면(예 '공동주택'이 여러 행) **첫 행에만 값을 쓰고 그 아래 반복 칸은 빈 문자열("")** 로 두세요(시스템이 세로 병합). 가로도 동일 — 합계 라벨은 첫 칸만, 나머지 빈칸.
5. 원문에 없는 숫자·사실을 만들지 마세요(환각 절대 금지). 있는 내용을 그대로 옮기기만.
6. 페이지 번호('3 / 26', '- 5 -'), 회사 로고, 반복 머리말/꼬리말은 버리세요.
7. 표가 매우 커도 행을 생략하지 말고 전부 넣으세요(슬라이드 분할은 시스템이 함).
8. 첨부 이미지뿐인 페이지(사업계획승인서 등)는 intro 한 줄만, tables/bullets 비움.
순수 JSON만 출력하고 그 외 텍스트·설명은 출력하지 마세요."""


def structure_page(raw_text: str, page_num: int, *, use_cache: bool = True,
                   verify_numbers: bool = True) -> dict:
    """
    PDF 페이지 원문 → 구조화 dict. 실패 시 None.

    Returns (성공 시): {
        section_label, subtitle, intro, bullets[], tables[], source,
        "_usage": {...}, "_hallucinated": [...]
    }
    """
    text = (raw_text or "").strip()
    if len(text) < 15:
        return None

    user_prompt = f"[PDF 페이지 {page_num} 원문]\n{text}"
    res = call_claude(
        system_prompt=SYSTEM_PROMPT,
        user_prompt=user_prompt,
        slide_num=page_num,
        pdf_context=text,
        use_cache=use_cache,
        prompt_version=PROMPT_VERSION,
    )
    if not res.get("ok") or not res.get("data"):
        return None

    data = res["data"]
    data.setdefault("section_num", None)
    data.setdefault("subtitle", "")
    data.setdefault("intro", "")
    data.setdefault("bullets", [])
    data.setdefault("tables", [])
    data.setdefault("source", "")
    data["_usage"] = res.get("usage")

    if verify_numbers:
        chk = verify_numbers_in_pdf(
            {k: v for k, v in data.items() if not k.startswith("_")}, text)
        data["_hallucinated"] = chk.get("hallucinated_numbers", [])
    return data


# ─────────────────────────────────────────────
# 섹션 분류 fallback (LLM이 section_num 못 주거나 범위밖일 때)
# ─────────────────────────────────────────────
_SEC_KEYWORDS = {
    2: ["금융조건", "대출조건", "투자구조", "구조도", "기초자산", "기한이익",
        "트랜치", "tranche", "ltv", "financing", "금리", "수수료"],
    4: ["appendix", "별첨", "현장사진", "시공사", "관계사", "사업계획승인",
        "실시계획", "인허가", "양해각서", "부록"],
    3: ["담보", "토지", "사업개요", "건축개요", "분양", "사업일정", "에쿼티",
        "equity", "사업수지", "입지", "시장", "사례", "차주", "시행사", "재무"],
    1: ["executive", "summary", "사모사채", "딜", "거래개요", "투자포인트"],
}


def _guess_section(struct: dict) -> int:
    blob = " ".join([
        str(struct.get("subtitle", "")),
        str(struct.get("intro", "")),
        " ".join(struct.get("bullets", []) or []),
        " ".join(t.get("title", "") for t in (struct.get("tables") or [])),
    ]).lower()
    for sec in (2, 4, 3, 1):   # 우선순위
        if any(kw in blob for kw in _SEC_KEYWORDS[sec]):
            return sec
    return 3   # 기본: 본건 사업 개요


def _is_structure_page(p: dict) -> bool:
    """투자구조도 페이지인지(병합 대상에서 제외)."""
    st = p.get("_struct") or {}
    blob = " ".join([
        str(st.get("subtitle", "")),
        " ".join(t.get("title", "") for t in (st.get("tables") or [])),
    ])
    return ("구조도" in blob) or bool(p.get("_invest_image") or p.get("_invest_name"))


def _merge_section2_financing(pages: list, *, debug: bool = False) -> None:
    """원본 IM에서 '주요 금융조건'은 하나의 「구분|내용」 표이지만 PDF가 여러 페이지로
       나뉘어 LLM이 페이지별로 쪼갠다. 사람 제안서처럼 섹션2(투자구조도 제외)의 금융조건
       페이지들을 **하나의 연속 label_value 표 + grid 표들**로 병합한다(소제목 '기초자산 개요').

       - 모든 페이지의 label_value 표 rows를 이어붙여 단일 label_value 표 1개 생성(맨 앞)
       - grid 표(대출조건·Cash-In/Out 등)는 순서대로 뒤에 그대로 유지
       → build_structured_slide가 길면 같은 소제목으로 여러 장에 흘림(사람 제안서와 동일).
    """
    fin = [p for p in pages
           if p.get("_sec_int") == 2 and p.get("_struct") and not _is_structure_page(p)]
    if len(fin) < 2:
        return

    lv_rows, grids = [], []
    for p in fin:
        for t in (p["_struct"].get("tables") or []):
            if (t.get("kind") or "") == "label_value":
                lv_rows.extend(t.get("rows") or [])
            else:
                grids.append(t)

    # ★사업명/시행사·차주 행을 표 맨 위에(원본: 구조도 밑부터 기초자산 표 = 사업명부터 시작).
    #   구조도 페이지(fin에서 제외됨)의 label_value 표에서 값을 가져와 신탁사 앞에 prepend.
    biz_top, _seen = [], set()
    for p in pages:
        for t in ((p.get("_struct") or {}).get("tables") or []):
            if (t.get("kind") or "") != "label_value":
                continue
            for r in (t.get("rows") or []):
                lab = str(r[0] if r else "").strip()
                val = str(r[1] if r and len(r) > 1 else "").strip()
                if not val:
                    continue
                if lab == "사업명" and "사업명" not in _seen:
                    biz_top.append(["사업명", val]); _seen.add("사업명")
                elif "시행사" in lab and "차주" in lab and "시행사" not in _seen:
                    biz_top.append([lab, val]); _seen.add("시행사")
    if biz_top:
        biz_top.sort(key=lambda r: 0 if r[0] == "사업명" else 1)
        # 이미 lv_rows에 같은 라벨 있으면 중복 제거
        _exist = {str(r[0]).strip() for r in lv_rows if r}
        biz_top = [r for r in biz_top if r[0] not in _exist]
        lv_rows = biz_top + lv_rows

    # ── ★표 안의 표: grid(대출조건·Cash 등)를 「구분|내용」 표의 해당 '행 셀' 안에 중첩 ──
    #   행 라벨(앵커)에 grid를 매달고, 렌더러가 그 행의 내용칸 위에 grid를 얹는다.
    def _find_row(sub):
        for i, r in enumerate(lv_rows):
            if sub in str(r[0] if r else ""):
                return i
        return None

    # 병합 페이지들의 각주(주N)) 수집 — grid별로 배정하고 남은 건 표 밑 일반 각주로
    all_notes = []
    for p in fin:
        for b in (p["_struct"].get("bullets") or []):
            if str(b).strip() and b not in all_notes:
                all_notes.append(str(b).strip())
        # ★LLM이 bullets로 안 뺀 경우 대비 — raw_text 에서 '주N) <내용>' 각주 정의 직접 추출
        #   (탁상감정가 980억…, 대여금 상환 방법… 등이 누락되던 문제)
        for ln in (p.get("raw_text", "") or "").splitlines():
            s = ln.strip()
            if s.startswith("주") and ")" in s:
                head = s[:s.index(")")].replace("주", "").strip()
                after = s[s.index(")") + 1:].strip()
                if head.isdigit() and len(after) >= 3 and s not in all_notes:
                    all_notes.append(s)

    def _take_note(keys):
        for i, nt in enumerate(all_notes):
            if any(k in nt for k in keys):
                return all_notes.pop(i)
        return ""

    nested = []   # [(anchor_label, grid_tdef, note), ...] note=그 grid 바로 밑에 붙일 주N)
    for g in grids:
        title = (g.get("title") or "")
        hdr = " ".join(str(h) for h in (g.get("header") or []))
        blob = title + " " + hdr
        if any(k in blob for k in ("Cash", "cash", "자금", "Cash-In", "Cash-Out")):
            anchor = "자금용도"
            note = _take_note(("대여금", "정산 방법", "기투입비용 정산 방"))
        else:
            anchor = "주요 대출조건"   # 트랜치별 대출금액·금리·LTV 등
            note = _take_note(("탁상감정", "감정가", "감정평가"))
        idx = _find_row(anchor)
        if idx is None:
            if anchor == "주요 대출조건":
                at = _find_row("대출기간")
                lv_rows.insert(at if at is not None else len(lv_rows), [anchor, ""])
            elif _find_row("자금용도") is None:
                lv_rows.append([anchor, ""])
        nested.append((anchor, g, note))

    merged_tables = []
    if lv_rows:
        _lvt = {"title": "", "kind": "label_value", "header": [], "rows": lv_rows}
        if all_notes:    # '본 금융조건은 당사자들간…' 등 남은 각주 → 표 바로 밑 9pt
            _lvt["_notes"] = list(all_notes)
        merged_tables.append(_lvt)

    base = fin[0]
    # 빨간 글씨/밑줄 합치기(병합된 페이지들 전부)
    reds, uls, fills = [], [], []
    for p in fin:
        reds.extend(p.get("_red_texts") or [])
        uls.extend(p.get("_underline_texts") or [])
        fills.extend(p.get("_filled_texts") or [])
    base["_red_texts"] = list(dict.fromkeys(reds))
    base["_underline_texts"] = list(dict.fromkeys(uls))
    base["_filled_texts"] = list(dict.fromkeys(fills))
    base["_struct"]["tables"] = merged_tables
    base["_struct"]["_nested_grids"] = nested
    base["_struct"]["subtitle"] = "기초자산 개요"
    base["_struct"]["bullets"] = all_notes   # grid에 안 붙은 일반 각주 → 표 바로 밑
    base["_struct"]["source"] = ""           # ★대전 원본엔 출처 없음 → 출처 제거
    # 나머지 금융 페이지는 제거(병합됨)
    drop = set(id(p) for p in fin[1:])
    pages[:] = [p for p in pages if id(p) not in drop]
    if debug:
        print(f"  [병합] 금융조건 {len(fin)}페이지 → 1개 표"
              f"(label_value {len(lv_rows)}행 + grid {len(grids)}개)")


def enrich_and_number(pages: list, *, debug: bool = False, pdf_path: str = None) -> list:
    """
    각 페이지를 LLM으로 구조화하고, 레인필드 4섹션 체계에 맞춰
    섹션라벨(고정) + 소제목 번호(x.y)를 부여해 page dict에 채워 넣는다.

    채워 넣는 키:
      page["_struct"]       : LLM 구조화 결과(dict) 또는 None(폴백 대상)
      page["section_num"]   : "01"~"04"
      page["section_name"]  : "섹션명"            (divider 제목용 — 번호 없음)
      page["section_title"] : "0N {섹션명}"        (TOC/섹션 인식 로직 호환)
      page["section_label"] : "0N {섹션명}"        (본문 슬라이드 좌상단)
      page["subtitle"]      : "N.k {소제목}"

    ★ 페이지를 섹션(1~4) 순서로 안정정렬한 뒤 번호를 부여한다.
      → divider가 섹션마다 한 번씩, 01~04 순서로만 생성됨(중복·뒤죽박죽 방지).
    """
    # ── 1패스: 각 페이지 구조화 + 섹션 분류 ──
    for idx, p in enumerate(pages):
        raw = p.get("raw_text", "") or ""
        st = structure_page(raw, p.get("page_num", 0)) if raw.strip() else None
        p["_struct"] = st
        if st:
            sec = st.get("section_num")
            try:
                sec = int(sec)
            except (TypeError, ValueError):
                sec = None
            if sec not in (1, 2, 3, 4):
                sec = _guess_section(st)
        else:
            sec = 4   # 구조화 실패(이미지 첨부 등) → Appendix
        p["_sec_int"] = sec
        p["_orig_idx"] = idx

    # ── 섹션 순서로 안정정렬 (섹션 내부는 원본 순서 유지) ──
    pages.sort(key=lambda p: (p.get("_sec_int", 4), p.get("_orig_idx", 0)))

    # ── 금융조건 = 하나의 표(여러 페이지로 쪼개진 것을 다시 합침) ──
    _merge_section2_financing(pages, debug=debug)

    # ── 각주 복원: raw_text의 '주N)' 각주를 살리고(글머리 X), '출처:' 아닌 표각주는 표 밑으로 ──
    _restore_page_notes(pages)

    # ── 투자구조도(2.1): 도형으로 직접 그림. 섹션2 맨 앞에 삽입 → 금융조건은 2.2로 밀림 ──
    try:
        from modules.ai_slide_builders import generate_investment_structure
        # 참여기관·트랜치 정보가 있는 섹션1·2 원문을 모아 데이터 생성
        src = "\n".join(p.get("raw_text", "") or "" for p in pages
                        if p.get("_sec_int") in (1, 2))[:6000]
        data, intro = {}, ""
        if src.strip():
            res = generate_investment_structure(src)
            if res.get("ok") and res.get("data"):
                data = res["data"]
                intro = (data.get("intro_paragraph") or "").strip()
        synth = {"_invest_diagram": data, "_invest_intro": intro,
                 "_invest_name": "투자구조도", "_sec_int": 2, "_struct": None,
                 "raw_text": "", "page_num": -1, "tables": [], "images": []}
        insert_at = next((i for i, p in enumerate(pages)
                          if p.get("_sec_int", 4) == 2), len(pages))
        pages.insert(insert_at, synth)
        if debug:
            print(f"  [투자구조도] 2.1 다이어그램 페이지 삽입(섹션2 맨 앞)")
    except Exception as _exc:
        print(f"[enrich] 투자구조도 생성 실패: {_exc}")

    # ── 2패스: 정렬된 순서로 x.y 번호 부여 ──
    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    for p in pages:
        sec = p.get("_sec_int", 4)
        counters[sec] += 1
        k = counters[sec]
        st = p.get("_struct")
        name = (p.get("_invest_name")
                or ((st.get("subtitle") if st else "") or "").strip()
                or SECTION_NAMES[sec])
        p["section_num"] = f"0{sec}"
        p["section_name"] = SECTION_NAMES[sec]
        p["section_title"] = f"0{sec} {SECTION_NAMES[sec]}"
        p["section_label"] = f"0{sec} {SECTION_NAMES[sec]}"
        p["subtitle"] = f"{sec}.{k} {name}"
        if st is not None:
            st["_final_subtitle"] = p["subtitle"]
            st["_final_section_label"] = p["section_label"]
        if debug:
            print(f"  p{p.get('page_num')}: {p['section_label']} | {p['subtitle']}")

    # ── 짧은 페이지 합치기(한 슬라이드에 들어갈 만한 연속 섹션3/4 페이지) ──
    _pack_short_pages(pages, debug=debug)
    return pages


def _restore_page_notes(pages: list) -> None:
    """각 페이지 각주 정리:
       ① raw_text의 '주N) …' 각주를 살려, prefix 없이 본문불릿(글머리 •)으로 들어간 것을 복원
          (예 '향후 사업 일정은…' → '주 1) 향후 사업 일정은…' → 9pt 각주로 렌더).
       ② '출처:' 형식이 아닌데 source에 들어간 표 각주(예 '금융감독원 전자공시시스템…')는
          해당 표(_notes)로 옮겨 표 바로 밑 9pt 로 — 출처(좌하단)로 안 가게."""
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        raw = p.get("raw_text", "") or ""
        raw_notes = []
        for ln in raw.splitlines():
            s = ln.strip()
            if s.startswith("주") and ")" in s[:6]:
                head = s[:s.index(")")].replace("주", "").strip()
                after = s[s.index(")") + 1:].strip()
                if head.isdigit() and len(after) >= 3:
                    raw_notes.append(s)

        def _match_raw(text):
            t = str(text).strip()
            for rn in raw_notes:
                body = rn[rn.index(")") + 1:].strip()
                if t and (t[:12] in body or body[:12] in t):
                    return rn
            return None

        # ① prefix 없는 각주 불릿 복원
        if raw_notes and st.get("bullets"):
            nb = []
            for b in st["bullets"]:
                bs = str(b).strip()
                nb.append(_match_raw(bs) if (not bs.startswith("주") and _match_raw(bs)) else b)
            st["bullets"] = nb

        # ② 출처가 '출처:' 형식이 아니면(표 각주) → 해당 표 _notes로
        src = (st.get("source") or "").strip()
        if src and ("출처" not in src):
            note_txt = _match_raw(src) or (src if src.startswith("주") else "주1) " + src)
            tabs = st.get("tables") or []
            target = None
            if any(k in src for k in ("재무제표", "감독원", "전자공시", "기준")):
                for t in tabs:
                    hd = " ".join(str(h) for h in (t.get("header") or []))
                    rowtxt = " ".join(str(c) for r in (t.get("rows") or []) for c in r)
                    if "회계연도" in hd or "회계연도" in rowtxt or "자산총계" in rowtxt:
                        target = t
                        break
            if target is None and tabs:
                target = tabs[-1]
            if target is not None:
                target.setdefault("_notes", []).append(note_txt)
                st["source"] = ""


def _pack_short_pages(pages: list, *, debug: bool = False) -> None:
    """연속된 섹션3/4 '짧은' 페이지를 한 슬라이드로 합쳐 페이지 수를 줄인다.
       (build_structured_slide가 넘치면 자동 분할하므로 합쳐도 안전 — 보수적 임계.)
       합칠 때: 첫 페이지 소제목이 슬라이드 상단, 뒤 페이지는 그 표에 미니라벨(소제목)로."""
    def can_lead(p):
        st = p.get("_struct")
        return (isinstance(st, dict) and p.get("_sec_int") in (3, 4)
                and not p.get("_invest_diagram") and not st.get("_nested_grids"))

    def est_h(st):
        h = 0.55 if (st.get("intro") or "").strip() else 0.12
        for t in (st.get("tables") or []):
            rows = t.get("rows") or []
            h += 0.40 + (1 + len(rows)) * 0.26
        h += 0.20 * len([b for b in (st.get("bullets") or []) if str(b).strip()])
        return h

    def can_absorb(p):
        # 뒤따르며 흡수될 작은 페이지(이미지 없음, 표 작음, 연락처/첨부이미지 제외)
        st = p.get("_struct")
        if not (isinstance(st, dict) and not p.get("_invest_diagram")
                and not st.get("_nested_grids") and not p.get("images")):
            return False
        sub = str(p.get("subtitle", ""))
        if any(k in sub for k in ("연락처", "담당자", "승인서", "양해각서", "공문")):
            return False
        return bool(st.get("tables")) and est_h(st) <= 2.7

    out, i = [], 0
    while i < len(pages):
        p = pages[i]
        # 리드 자격 + '표가 있는' 페이지만 뒤 페이지를 흡수(이미지 전용/연락처는 흡수 안 함)
        lead_sub = str(p.get("subtitle", ""))
        lead_ok = (can_lead(p) and bool(p["_struct"].get("tables"))
                   and not any(k in lead_sub for k in ("연락처", "담당자", "승인서", "양해각서", "공문")))
        if not lead_ok:
            out.append(p); i += 1; continue
        group, j = [p], i + 1
        # 큰 페이지(p)가 리드 → 뒤따르는 같은 섹션의 '작은' 페이지들을 흡수
        while (j < len(pages) and pages[j].get("_sec_int") == p.get("_sec_int")
               and can_absorb(pages[j])):
            group.append(pages[j]); j += 1
        if len(group) > 1:
            base = group[0]; bst = base["_struct"]

            def _is_note_b(b):
                s = str(b).lstrip()
                return s.startswith(("주", "*", "※")) and (
                    s[:1] in ("*", "※") or (len(s) > 1 and (s[1].isdigit() or s[1] == " ")))

            # ★base(앞 표)의 각주는 base의 '마지막 표'에 귀속(아래 표로 안 넘어가게)
            _base_notes = [b for b in (bst.get("bullets") or []) if _is_note_b(b)]
            bst["bullets"] = [b for b in (bst.get("bullets") or []) if not _is_note_b(b)]
            if _base_notes and bst.get("tables"):
                bst["tables"][-1].setdefault("_notes", []).extend(_base_notes)

            for q in group[1:]:
                qst = q["_struct"]
                qname = q.get("subtitle", "")          # "3.2 ..."
                qtabs = qst.get("tables") or []
                # 이 페이지(q)의 각주는 이 페이지 '첫 표' 밑에만 붙인다(표별 각주 분리)
                q_notes = [b for b in (qst.get("bullets") or []) if _is_note_b(b)]
                q_prose = [b for b in (qst.get("bullets") or []) if not _is_note_b(b)]
                for k, t in enumerate(qtabs):
                    t2 = dict(t)
                    if k == 0 and not (t2.get("title") or "").strip():
                        t2["title"] = qname             # 뒤 페이지 첫 표에 소제목 라벨
                    if k == 0 and q_notes:
                        t2.setdefault("_notes", []).extend(q_notes)
                    bst.setdefault("tables", []).append(t2)
                # ★흡수된 페이지의 intro(LLM 생성 소개문, 원본엔 없음)는 떠다니는 불릿으로
                #   찍지 않는다(사용자: "적지마" — 표만 깔끔하게). 표 옆 산문(분석)만 유지.
                bst.setdefault("bullets", []).extend(q_prose)
            out.append(base)
            if debug:
                print(f"  [페이지합치기] {[g.get('subtitle') for g in group]} → 1슬라이드")
        else:
            out.append(p)
        i = j
    pages[:] = out
