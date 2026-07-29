# MIS 마우스 좌클릭 좌표 기록기
"""
좌클릭마다 (절대 좌표, MIS 창 기준 상대 좌표)를 콘솔과 로그 파일에 기록한다.
tools/AI-Elite-MIS_RPA/production_coords.json / utility_coords.json 작성 시 사용.

사용법:
    1. MIS (신)종합정보 실행 + 대상 화면 진입
    2. python mouse_click_logger.py
    3. 측정할 UI 요소를 차례로 좌클릭
    4. Ctrl+C 종료

특징:
    - 좌버튼 다운 트랜지션 1회만 기록 (디바운싱 자동)
    - 0.3초 + 5px 이내 두 번째 클릭은 [DBL] 더블클릭으로 묶어 기록
    - MIS 창 영역 밖 클릭은 [외부] 마커 표시
    - 로그 파일: logs/click_log_YYYYMMDD_HHMMSS.txt
"""

# 더블클릭 판정 임계값
DOUBLE_CLICK_TIME = 0.3       # 초
DOUBLE_CLICK_DISTANCE = 5     # 픽셀

import ctypes
import time
from datetime import datetime
from pathlib import Path

from pywinauto import Application

VK_LBUTTON = 0x01


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


def get_cursor_pos():
    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def is_lbutton_pressed():
    # GetAsyncKeyState: 최상위 비트가 1이면 키가 눌린 상태
    return bool(ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000)


def main():
    print("=== MIS 마우스 좌클릭 좌표 기록기 ===")
    try:
        app = Application(backend="uia").connect(title_re=".*종합정보.*")
        win = app.window(title_re=".*종합정보.*")
        rect = win.rectangle()
        print(f"MIS 창 탐지: ({rect.left}, {rect.top}) ~ ({rect.right}, {rect.bottom})")
        print(f"  크기: {rect.width()} x {rect.height()}")
    except Exception as e:
        print(f"[오류] MIS 창을 찾을 수 없습니다: {e}")
        print("MIS (신)종합정보를 먼저 실행해주세요.")
        return

    log_dir = Path(__file__).parent / "logs"
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / f"click_log_{datetime.now():%Y%m%d_%H%M%S}.txt"

    print(f"로그 파일: {log_path}")
    print("-" * 70)
    print("좌클릭마다 좌표가 기록됩니다. 종료: Ctrl+C")
    print("-" * 70)

    counter = 0
    prev_pressed = False
    pending = None   # (x, y, rel_x, rel_y, t, in_window) — 단일/더블 판정 보류 중인 클릭

    def write(line: str):
        print(line)
        flog.write(line + "\n")
        flog.flush()

    def format_line(idx, t_str, x, y, rel_x, rel_y, in_window, dbl=False):
        tags = []
        if dbl:
            tags.append("[DBL]")
        if not in_window:
            tags.append("[외부]")
        suffix = ("  " + " ".join(tags)) if tags else ""
        return (f"[{idx:02d}] {t_str}  "
                f"abs({x:>5}, {y:>5})  rel({rel_x:>5}, {rel_y:>5}){suffix}")

    def flush_pending():
        """pending이 더블클릭 짝을 만나지 못한 경우 단일 클릭으로 확정 기록."""
        nonlocal pending, counter
        if pending is None:
            return
        x, y, rel_x, rel_y, t, in_win = pending
        counter += 1
        t_str = datetime.fromtimestamp(t).strftime("%H:%M:%S.%f")[:-3]
        write(format_line(counter, t_str, x, y, rel_x, rel_y, in_win, dbl=False))
        pending = None

    with open(log_path, "w", encoding="utf-8") as flog:
        flog.write(f"MIS 마우스 좌클릭 좌표 기록 — 시작 {datetime.now()}\n")
        flog.write(f"창 좌상단: ({rect.left}, {rect.top})  크기: "
                   f"{rect.width()}x{rect.height()}\n")
        flog.write(f"더블클릭 임계값: {DOUBLE_CLICK_TIME}s, {DOUBLE_CLICK_DISTANCE}px\n")
        flog.write("=" * 70 + "\n")
        flog.flush()

        try:
            while True:
                pressed = is_lbutton_pressed()

                # 다운 트랜지션(0→1) — 클릭 발생
                if pressed and not prev_pressed:
                    x, y = get_cursor_pos()
                    rel_x = x - rect.left
                    rel_y = y - rect.top
                    now = time.time()
                    in_window = (rect.left <= x <= rect.right and
                                 rect.top <= y <= rect.bottom)

                    if pending is not None:
                        px, py, _, _, pt, p_in = pending
                        if (now - pt <= DOUBLE_CLICK_TIME and
                                abs(x - px) <= DOUBLE_CLICK_DISTANCE and
                                abs(y - py) <= DOUBLE_CLICK_DISTANCE):
                            # 더블클릭 확정 — 첫 클릭 좌표/시각 기준으로 1줄 기록
                            counter += 1
                            t_str = datetime.fromtimestamp(pt).strftime("%H:%M:%S.%f")[:-3]
                            prx, pry = pending[2], pending[3]
                            write(format_line(counter, t_str, px, py,
                                              prx, pry, p_in, dbl=True))
                            pending = None
                            prev_pressed = pressed
                            time.sleep(0.01)
                            continue
                        else:
                            # 시간/거리 초과 — 이전 pending은 단일 클릭으로 확정
                            flush_pending()

                    # 새 pending 저장 (단일/더블 판정 대기)
                    pending = (x, y, rel_x, rel_y, now, in_window)

                # pending 시간 초과 시 단일 클릭으로 확정
                if pending is not None and (time.time() - pending[4]) > DOUBLE_CLICK_TIME:
                    flush_pending()

                prev_pressed = pressed
                time.sleep(0.01)

        except KeyboardInterrupt:
            flush_pending()
            footer = f"\n측정 종료. 총 {counter}회 클릭 기록."
            print(footer)
            flog.write(f"\n총 {counter}회 클릭 — 종료 {datetime.now()}\n")
            print(f"로그: {log_path}")


if __name__ == "__main__":
    main()
