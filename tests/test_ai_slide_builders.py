"""
tests/test_ai_slide_builders.py
────────────────────────────────────────────────────────
5-4-C: generate_*() 함수 통합 테스트
천안 부성2지구 PDF 원문을 Claude API 로 분석 → 결과 검증.
────────────────────────────────────────────────────────
"""

import sys
import os
import json
import traceback

# UTF-8 출력 강제 (Windows cp949 깨짐 방지)
sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)
sys.stderr = open(sys.stderr.fileno(), mode="w", encoding="utf-8", buffering=1)

# 프로젝트 루트를 sys.path 에 추가
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

PDF_PATH = os.path.join(ROOT, "★ TS_천안 부성2지구 도시개발사업 BL_v5.2.pdf")


# ─────────────────────────────────────────────
# PDF 텍스트 추출
# ─────────────────────────────────────────────
def extract_pdf_text(pdf_path: str) -> str:
    import fitz
    doc = fitz.open(pdf_path)
    pages = [page.get_text() for page in doc]
    doc.close()
    return "\n".join(pages)


# ─────────────────────────────────────────────
# 검증 행 출력 헬퍼
# ─────────────────────────────────────────────
def _status_icon(status: str) -> str:
    if status == "PASS":
        return "✓ PASS"
    if status == "FAIL":
        return "✗ FAIL"
    if status == "INFO":
        return "  INFO"
    if status == "SKIP":
        return "  SKIP"
    if "경고" in status:
        return f"⚠  {status}"
    return status


def print_row(code: str, content: str, status: str):
    icon = _status_icon(status)
    # 60자 초과 시 말줄임
    content_disp = (content[:57] + "...") if len(content) > 60 else content
    print(f"| {code:<3} | {content_disp:<60} | {icon:<10} |")


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def main():
    print("=" * 80)
    print("5-4-C: AI 슬라이드 빌더 통합 테스트 (천안 부성2지구 PDF)")
    print("=" * 80)

    # ── API 키 확인 ──────────────────────────────────────────
    from modules.claude_api import get_client, verify_numbers_in_pdf, estimate_cost
    try:
        get_client()
        print("[OK] API 클라이언트 초기화 완료")
    except Exception as exc:
        print(f"[FAIL] API 키 없음 또는 초기화 실패: {exc}")
        print("       .env 파일에 ANTHROPIC_API_KEY 를 설정하세요.")
        sys.exit(1)

    # ── PDF 존재 확인 ─────────────────────────────────────────
    if not os.path.exists(PDF_PATH):
        print(f"[FAIL] PDF 파일 없음: {PDF_PATH}")
        sys.exit(1)

    # ── 1. PDF 텍스트 추출 ────────────────────────────────────
    print(f"\n[STEP 1] PDF 텍스트 추출")
    print(f"  경로: {PDF_PATH}")
    try:
        pdf_text = extract_pdf_text(PDF_PATH)
        print(f"  추출 완료 — {len(pdf_text):,} 자")
    except Exception:
        print(f"[FAIL] PDF 추출 실패:\n{traceback.format_exc()}")
        sys.exit(1)

    # ── 2. 3개 generate 함수 호출 ─────────────────────────────
    from modules.ai_slide_builders import (
        generate_executive_summary,
        generate_sasae_overview,
        generate_investment_structure,
    )

    results: dict = {}
    total_input = 0
    total_output = 0

    calls = [
        ("exec_summary",  "슬라이드 2 — Executive Summary",     generate_executive_summary),
        ("sasae_overview","슬라이드 5 — 사모사채 개요",          generate_sasae_overview),
        ("investment",    "슬라이드 7 — 투자구조도",             generate_investment_structure),
    ]

    print(f"\n[STEP 2] Claude API 호출 (총 {len(calls)}회)")
    for key, label, fn in calls:
        print(f"\n  >>> {label}")
        try:
            r = fn(pdf_text)
            results[key] = r
            if r.get("ok"):
                u = r["usage"]
                total_input  += u.get("input_tokens",  0)
                total_output += u.get("output_tokens", 0)
                cache_str = "캐시 HIT" if r.get("cached") else "신규 API"
                print(f"      OK | {cache_str} | in={u['input_tokens']:,} out={u['output_tokens']:,} | ${u.get('estimated_cost_usd',0):.4f}")
            else:
                print(f"      FAIL — {r.get('error')}")
        except Exception:
            results[key] = {"ok": False, "data": {}, "error": traceback.format_exc()}
            print(f"      EXCEPTION:\n{traceback.format_exc()}")

    # ── 3. 검증 ──────────────────────────────────────────────
    print("\n" + "=" * 80)
    print("검증 결과")
    print("=" * 80)
    print(f"| {'항목':<3} | {'검증 내용':<60} | {'결과':<10} |")
    print(f"|{'-'*5}|{'-'*62}|{'-'*12}|")

    all_rows: list = []

    # ── Executive Summary (A) ──────────────────────────────
    r2 = results.get("exec_summary", {})
    d2 = r2.get("data", {}) if r2.get("ok") else {}
    kps = d2.get("key_points", [])

    a1 = len(kps) == 5
    all_rows.append(("A1", f"key_points 5개 (실제 {len(kps)}개)", "PASS" if a1 else "FAIL"))

    a2 = all("title" in kp and "description" in kp for kp in kps) if kps else False
    all_rows.append(("A2", "각 항목에 title + description 키", "PASS" if a2 else "FAIL"))

    titles_str = ", ".join(kp.get("title", "?") for kp in kps)
    all_rows.append(("A3", f"titles: {titles_str}", "INFO"))

    if d2 and r2.get("ok"):
        v2 = verify_numbers_in_pdf(d2, pdf_text)
        hall2 = v2["hallucinated_numbers"]
        a4_st = "PASS" if v2["ok"] else f"경고 {len(hall2)}개"
        hall2_str = str(hall2[:3]) if hall2 else "없음"
        all_rows.append(("A4", f"환각 의심 숫자: {hall2_str}", a4_st))
    else:
        all_rows.append(("A4", "API 실패로 검증 불가", "SKIP"))

    # ── 사모사채 개요 (B) ─────────────────────────────────
    r5 = results.get("sasae_overview", {})
    d5 = r5.get("data", {}) if r5.get("ok") else {}
    fields = d5.get("fields", [])

    b1 = len(fields) == 9
    all_rows.append(("B1", f"fields 9개 (실제 {len(fields)}개)", "PASS" if b1 else "FAIL"))

    issuing_amount = next(
        (f.get("value", "") for f in fields if f.get("label") == "발행금액"), ""
    )
    b2 = "440" in issuing_amount
    all_rows.append(("B2", f"발행금액에 '440' 포함: '{issuing_amount[:40]}'", "PASS" if b2 else "FAIL"))

    if d5 and r5.get("ok"):
        v5 = verify_numbers_in_pdf(d5, pdf_text)
        hall5 = v5["hallucinated_numbers"]
        b3_st = "PASS" if v5["ok"] else f"경고 {len(hall5)}개"
        hall5_str = str(hall5[:3]) if hall5 else "없음"
        all_rows.append(("B3", f"환각 의심 숫자: {hall5_str}", b3_st))
    else:
        all_rows.append(("B3", "API 실패로 검증 불가", "SKIP"))

    # ── 투자구조도 (C) ────────────────────────────────────
    r7 = results.get("investment", {})
    d7 = r7.get("data", {}) if r7.get("ok") else {}

    req_keys = {"total_loan_amount", "tranches", "entities", "relationships"}
    c1 = req_keys.issubset(set(d7.keys())) if d7 else False
    missing = req_keys - set(d7.keys()) if d7 else req_keys
    all_rows.append(("C1", f"필수 키 4개 (누락: {missing or '없음'})", "PASS" if c1 else "FAIL"))

    tranches = d7.get("tranches", [])
    c2 = len(tranches) == 2
    all_rows.append(("C2", f"tranches 2개 (실제 {len(tranches)}개)", "PASS" if c2 else "FAIL"))

    borrower = d7.get("entities", {}).get("borrower", "")
    c3 = "더함도시개발" in borrower
    all_rows.append(("C3", f"borrower 확인: '{borrower[:40]}'", "PASS" if c3 else "FAIL"))

    if d7 and r7.get("ok"):
        v7 = verify_numbers_in_pdf(d7, pdf_text)
        hall7 = v7["hallucinated_numbers"]
        c4_st = "PASS" if v7["ok"] else f"경고 {len(hall7)}개"
        hall7_str = str(hall7[:3]) if hall7 else "없음"
        all_rows.append(("C4", f"환각 의심 숫자: {hall7_str}", c4_st))
    else:
        all_rows.append(("C4", "API 실패로 검증 불가", "SKIP"))

    for code, content, status in all_rows:
        print_row(code, content, status)

    # ── 4. Raw JSON ───────────────────────────────────────
    print("\n" + "=" * 80)
    print("RAW JSON 응답")
    print("=" * 80)

    raw_labels = [
        ("Executive Summary (슬라이드 2)", "exec_summary"),
        ("사모사채 개요 (슬라이드 5)",     "sasae_overview"),
        ("투자구조도 (슬라이드 7)",         "investment"),
    ]
    for label, key in raw_labels:
        r = results.get(key, {})
        print(f"\n── {label} ──")
        if r.get("ok"):
            print(json.dumps(r["data"], ensure_ascii=False, indent=2))
        else:
            err = r.get("error", "알 수 없는 오류")
            print(f"  [FAIL] {err}")

    # ── 5. 토큰 / 비용 요약 ───────────────────────────────
    total_cost = estimate_cost(total_input, total_output)
    print("\n" + "=" * 80)
    print("토큰 / 비용 요약")
    print("=" * 80)
    print(f"  총 input  tokens : {total_input:,}")
    print(f"  총 output tokens : {total_output:,}")
    print(f"  추정 비용        : ${total_cost:.4f} USD")

    # ── 6. 캐시 검증 ─────────────────────────────────────
    print("\n" + "=" * 80)
    print("캐시 검증 (Executive Summary 동일 입력 재호출)")
    print("=" * 80)
    try:
        r_cache = generate_executive_summary(pdf_text)
        if r_cache.get("cached"):
            print("  결과: 캐시 HIT ✓  (두 번째 호출에서 캐시 반환 확인)")
        else:
            print("  결과: 캐시 MISS ✗  (캐시가 작동하지 않거나 TTL 만료)")
    except Exception:
        print(f"  예외:\n{traceback.format_exc()}")

    # ── 최종 요약 ─────────────────────────────────────────
    pass_n = sum(1 for _, _, s in all_rows if s == "PASS")
    fail_n = sum(1 for _, _, s in all_rows if s == "FAIL")
    warn_n = sum(1 for _, _, s in all_rows if "경고" in s)
    info_n = sum(1 for _, _, s in all_rows if s in ("INFO", "SKIP"))

    print(f"\n{'='*80}")
    print(f"최종 결과:  {pass_n} PASS  /  {fail_n} FAIL  /  {warn_n} 경고  /  {info_n} INFO·SKIP")
    print(f"{'='*80}")


if __name__ == "__main__":
    main()
