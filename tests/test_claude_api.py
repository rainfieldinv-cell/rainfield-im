"""
tests/test_claude_api.py
─────────────────────────────────────────────────────────
modules/claude_api.py 기능 검증 스크립트.

실행 방법:
  python tests/test_claude_api.py
─────────────────────────────────────────────────────────
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from modules.claude_api import (
    get_client,
    call_claude,
    verify_numbers_in_pdf,
    estimate_cost,
    clear_cache,
    get_cache_stats,
)

PASS = "[PASS]"
FAIL = "[FAIL]"
results = []


def check(label: str, condition: bool, detail: str = ""):
    mark = PASS if condition else FAIL
    msg  = f"  {mark} {label}" + (f" - {detail}" if detail else "")
    print(msg)
    results.append((label, condition))
    return condition


# ══════════════════════════════════════════════════════════
# TEST 1: get_client() - API 키 로드
# ══════════════════════════════════════════════════════════
print("\n[TEST 1] get_client() - API 키 로드")
try:
    client = get_client()
    check("API 키 로드 성공", True, f"client type={type(client).__name__}")
except ValueError as e:
    check("API 키 로드 성공", False, str(e))
    print("\n  !! ANTHROPIC_API_KEY 미설정. 테스트 2 (실제 API 호출)를 건너뜁니다.")
    client = None
except Exception as e:
    check("API 키 로드 성공", False, str(e))
    client = None


# ══════════════════════════════════════════════════════════
# TEST 2: call_claude() - 더미 호출
# ══════════════════════════════════════════════════════════
print("\n[TEST 2] call_claude() - 더미 API 호출")

DUMMY_SYSTEM = (
    "당신은 PDF 분석가입니다. "
    "사용자가 주는 텍스트에서 핵심 숫자를 모두 찾아 "
    '{"numbers": ["숫자1", "숫자2", ...]} 형식의 JSON으로만 반환하세요.'
)
DUMMY_USER = (
    "본건 대출금액은 1,640억원이며, Tr.A 1,200억원, Tr.B 440억원으로 구성된다."
)
DUMMY_PDF_CONTEXT = DUMMY_USER

res = None
if client is not None:
    clear_cache()   # 테스트 전 캐시 초기화
    res = call_claude(
        system_prompt=DUMMY_SYSTEM,
        user_prompt=DUMMY_USER,
        slide_num=0,
        pdf_context=DUMMY_PDF_CONTEXT,
        use_cache=True,
    )
    check("ok=True",       res["ok"],       f"error={res.get('error')}")
    check("cached=False",  not res["cached"], "첫 호출은 캐시 미스여야 함")
    check("data dict 반환", isinstance(res.get("data"), dict), str(res.get("data")))

    numbers = res.get("data", {}).get("numbers", [])
    has_1640 = any("1640" in str(n).replace(",", "") for n in numbers)
    check("1,640 포함",    has_1640,    f"numbers={numbers}")

    usage = res.get("usage", {})
    check("input_tokens > 0",  usage.get("input_tokens", 0) > 0,
          f"input={usage.get('input_tokens')}")
    check("output_tokens > 0", usage.get("output_tokens", 0) > 0,
          f"output={usage.get('output_tokens')}")
    check("cost > 0",          usage.get("estimated_cost_usd", 0) > 0,
          f"${usage.get('estimated_cost_usd', 0):.5f}")
    print(f"\n  실제 사용 토큰: input={usage.get('input_tokens')} / output={usage.get('output_tokens')}")
    print(f"  추정 비용: ${usage.get('estimated_cost_usd', 0):.5f} USD")
else:
    print("  (건너뜀 - API 키 없음)")


# ══════════════════════════════════════════════════════════
# TEST 3: verify_numbers_in_pdf()
# ══════════════════════════════════════════════════════════
print("\n[TEST 3] verify_numbers_in_pdf() - 숫자 환각 검증")

PDF_SAMPLE = (
    "본건 대출금액은 1,640억원이며, Tr.A 1,200억원, Tr.B 440억원으로 구성된다. "
    "LTV 81.48%, 이자율 8.72%, 만기 2029.09."
)

# 3-A: PDF에 있는 숫자만 → ok=True
good_output = {"numbers": ["1,640", "1,200", "440"], "ltv": "81.48%"}
v_good = verify_numbers_in_pdf(good_output, PDF_SAMPLE)
check("PDF 내 숫자 → ok=True",
      v_good["ok"],
      f"hallucinated={v_good['hallucinated_numbers']}")
check("verified_count == total_count",
      v_good["verified_count"] == v_good["total_count"],
      f"{v_good['verified_count']}/{v_good['total_count']}")

# 3-B: 가짜 숫자 포함 → hallucinated에 잡혀야 함
bad_output = {"numbers": ["1,640", "9999", "77777"], "fake_amount": "12345억"}
v_bad = verify_numbers_in_pdf(bad_output, PDF_SAMPLE)
check("가짜 숫자 → ok=False",
      not v_bad["ok"],
      f"hallucinated={v_bad['hallucinated_numbers']}")
check("hallucinated_numbers 비어있지 않음",
      len(v_bad["hallucinated_numbers"]) > 0,
      str(v_bad["hallucinated_numbers"]))


# ══════════════════════════════════════════════════════════
# TEST 4: 캐시 동작
# ══════════════════════════════════════════════════════════
print("\n[TEST 4] 캐시 동작 검증")

if client is not None and res is not None and res["ok"]:
    res2 = call_claude(
        system_prompt=DUMMY_SYSTEM,
        user_prompt=DUMMY_USER,
        slide_num=0,
        pdf_context=DUMMY_PDF_CONTEXT,
        use_cache=True,
    )
    check("두 번째 호출 cached=True", res2["cached"], "캐시 히트여야 함")
    check("캐시 결과 동일",
          res2.get("data") == res.get("data"),
          f"data={res2.get('data')}")

    stats = get_cache_stats()
    check("캐시 파일 1개 이상",
          stats["file_count"] >= 1,
          f"files={stats['file_count']} size={stats['total_mb']}MB")
else:
    print("  (건너뜀 - 이전 API 호출 없음)")


# ══════════════════════════════════════════════════════════
# TEST 5: estimate_cost() 단위 검증
# ══════════════════════════════════════════════════════════
print("\n[TEST 5] estimate_cost() 단위 검증")
cost_1m = estimate_cost(1_000_000, 0)
check("input 1M tokens = $3.00", abs(cost_1m - 3.0) < 0.001, f"${cost_1m}")
cost_out = estimate_cost(0, 1_000_000)
check("output 1M tokens = $15.00", abs(cost_out - 15.0) < 0.001, f"${cost_out}")


# ══════════════════════════════════════════════════════════
# 최종 요약
# ══════════════════════════════════════════════════════════
print("\n" + "=" * 50)
passed = sum(1 for _, ok in results if ok)
total  = len(results)
print(f"결과: {passed}/{total} 통과")
if passed == total:
    print("모든 테스트 통과!")
else:
    failed = [label for label, ok in results if not ok]
    print(f"실패 항목: {failed}")
print("=" * 50)
