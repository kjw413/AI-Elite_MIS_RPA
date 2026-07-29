# MIS 창 기준 좌표 측정 도구
"""
MIS (신)종합정보 창 안에서 마우스 위치의
절대 좌표와 창 기준 상대 좌표를 실시간으로 표시한다.

사용법:
    1. MIS (신)종합정보를 먼저 실행
    2. python mouse_pos.py
    3. 창 내 코드의 타겟 버튼/드롭다운으로 마우스 이동
    4. 읽은 상대 좌표를 JSON에 입력 (Ctrl+C 종료)
"""

import sys
import time
import ctypes

from pywinauto import Application


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


def get_cursor_pos():
    """Win32 API로 마우스 절대 좌표 조회 (의존성 없음)."""
    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def main():
    print("=== MIS 좌표 측정 도구 ===")
    try:
        app = Application(backend="uia").connect(title_re=".*종합정보.*")
        win = app.window(title_re=".*종합정보.*")
        rect = win.rectangle()
        print(f"MIS 창 탐지: {rect}")
        print(f"  좌상단: ({rect.left}, {rect.top})  크기: {rect.width()} x {rect.height()}")
    except Exception as e:
        print(f"[오류] MIS 창을 찾을 수 없습니다: {e}")
        print("MIS (신)종합정보를 먼저 실행해주세요.")
        return

    print("-" * 60)
    print("마우스를 창 내부로 이동하면 좌표가 표시됩니다.")
    print("JSON에 입력할 값은 '창 기준(상대)' 값입니다.")
    print("종료: Ctrl+C")
    print("-" * 60)

    try:
        while True:
            x, y = get_cursor_pos()
            rel_x = x - rect.left
            rel_y = y - rect.top
            sys.stdout.write(
                f"\r화면(절대): {x:>4}, {y:>4}  |  창 기준(상대): {rel_x:>4}, {rel_y:>4}    "
            )
            sys.stdout.flush()
            time.sleep(0.05)
    except KeyboardInterrupt:
        print("\n\n측정을 종료합니다.")


if __name__ == "__main__":
    main()
