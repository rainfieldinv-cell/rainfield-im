"""메모 저장 — 사업명/문제점/추가의견 텍스트를 로컬 JSON(memos.json)에 보관.

★PPT 파일은 저장하지 않는다(텍스트 메모만). 자동 삭제 없음 — 삭제는 사용자가 수동으로.
저장 위치: 저장소 루트의 memos.json.
"""
import os
import json

_MEMO_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "memos.json")


def load_memos():
    """저장된 메모 리스트 반환(없거나 읽기 실패 시 빈 리스트)."""
    try:
        with open(_MEMO_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_memos(memos):
    try:
        with open(_MEMO_PATH, "w", encoding="utf-8") as f:
            json.dump(memos, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def add_memo(business, problem, opinion, created):
    """새 메모를 맨 앞에 추가하고 전체 리스트 반환."""
    memos = load_memos()
    memos.insert(0, {
        "business": business,
        "problem": problem,
        "opinion": opinion,
        "created": created,
    })
    save_memos(memos)
    return memos


def delete_memo(index):
    """index 위치 메모 삭제 후 전체 리스트 반환."""
    memos = load_memos()
    if 0 <= index < len(memos):
        memos.pop(index)
        save_memos(memos)
    return memos
