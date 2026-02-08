import os
import time
import threading
import ctypes
from ctypes import wintypes

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# sorter.py에서 재사용할 것들 가져오기
from sorter import (
    load_config,
    ensure_dirs,
    log_line,
    safe_name,
    bucket_for_ext,
    wait_until_ready,
)


# -----------------------------
# Windows: 글로벌 핫키(F8) 등록
# -----------------------------
user32 = ctypes.windll.user32

WM_HOTKEY = 0x0312
HOTKEY_ID = 1
VK_F8 = 0x77  # F8
MOD_NOMOD = 0x0000


def get_foreground_window_title() -> str:
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return ""
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value or ""


def extract_room_from_title(title: str) -> str:
    """
    카카오톡 창 제목에서 채팅방 이름만 뽑아냅니다.
    일반적으로: "채팅방이름 - 카카오톡" 또는 "채팅방이름 - KakaoTalk"
    """
    t = (title or "").strip()
    if not t:
        return "미분류"

    # 흔한 접미사 제거
    for suffix in [" - 카카오톡", " - KakaoTalk", "- 카카오톡", "- KakaoTalk"]:
        if t.endswith(suffix):
            t = t[: -len(suffix)].strip()

    # 혹시 제목에 "카카오톡"만 남는 경우
    t = t.replace("카카오톡", "").replace("KakaoTalk", "").strip()
    if not t:
        return "미분류"

    # 파일/폴더 이름에 못 쓰는 문자 제거
    t = safe_name(t)
    return t if t else "미분류"


class Context:
    """
    F8을 눌렀을 때의 '채팅방 컨텍스트'를 메모리에 보관합니다.
    """
    def __init__(self):
        self.lock = threading.Lock()
        self.room = "미분류"
        self.ts_str = time.strftime("%Y%m%d%H%M%S")
        self.ts_epoch = 0.0  # 마지막 캡처 시각(epoch)

    def set(self, room: str):
        with self.lock:
            self.room = room
            self.ts_str = time.strftime("%Y%m%d%H%M%S")
            self.ts_epoch = time.time()

    def get(self, ttl_seconds: int) -> tuple[str, str]:
        with self.lock:
            age = time.time() - self.ts_epoch
            if self.ts_epoch > 0 and age <= ttl_seconds:
                return self.room, self.ts_str
        return "미분류", time.strftime("%Y%m%d%H%M%S")


def hotkey_thread_fn(ctx: Context, stop_event: threading.Event):
    """
    F8 글로벌 핫키를 등록하고 메시지 루프를 돌립니다.
    """
    # 기존 등록이 남아있을 수 있으니 해제 시도
    try:
        user32.UnregisterHotKey(None, HOTKEY_ID)
    except Exception:
        pass

    if not user32.RegisterHotKey(None, HOTKEY_ID, MOD_NOMOD, VK_F8):
        # 등록 실패 (다른 프로그램이 점유했거나 권한/환경 문제)
        print("ERROR: F8 핫키 등록 실패. 다른 프로그램에서 F8을 사용 중일 수 있어요.")
        print("      (추후 설정으로 핫키 변경 기능을 추가할 수 있습니다.)")
        return

    msg = wintypes.MSG()
    while not stop_event.is_set():
        # GetMessageW는 메시지가 올 때까지 대기합니다.
        # WM_HOTKEY를 받으면 room 컨텍스트를 갱신합니다.
        res = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        if res == 0:
            break  # WM_QUIT
        if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_ID:
            title = get_foreground_window_title()
            room = extract_room_from_title(title)
            ctx.set(room)
            # 사용자 피드백(콘솔)
            print(f"[F8] room captured: {room}")

    user32.UnregisterHotKey(None, HOTKEY_ID)


# -----------------------------
# 다운로드 폴더 감시 & 정리
# -----------------------------
class Handler(FileSystemEventHandler):
    def __init__(self, cfg, ctx: Context):
        super().__init__()
        self.cfg = cfg
        self.ctx = ctx

    def on_created(self, event):
        if event.is_directory:
            return

        src = event.src_path
        name = os.path.basename(src)
        _, ext = os.path.splitext(name)
        ext_l = ext.lower()

        if ext_l in self.cfg.ignore_ext:
            return

        if not wait_until_ready(src):
            log_line(self.cfg, f"SKIP not-ready: {src}")
            return

        room, ts = self.ctx.get(self.cfg.hotkey_context_ttl_seconds)

        # 파일명에는 "분까지만" 쓰기 (YYYYMMDDHHMM)
        ts = ts[:12]

        room = safe_name(room)
        bucket = bucket_for_ext(self.cfg, ext_l)
        orig_safe = safe_name(name)

        new_name = self.cfg.rename_template.format(
            ts=ts, room=room, bucket=bucket, orig=orig_safe
        )
        new_name = safe_name(new_name)

        dst_dir = os.path.join(self.cfg.output_dir, room, bucket)
        os.makedirs(dst_dir, exist_ok=True)

        dst = os.path.join(dst_dir, new_name)

        base, e2 = os.path.splitext(dst)
        i = 1
        while os.path.exists(dst):
            dst = f"{base}({i}){e2}"
            i += 1

        try:
            import shutil
            shutil.move(src, dst)
            log_line(self.cfg, f"MOVED {src} -> {dst}")
        except Exception as ex:
            log_line(self.cfg, f"FAIL move: {src} ({ex})")


def main():
    cfg = load_config()
    ensure_dirs(cfg)

    print("=== Kakao Download Organizer (single app) ===")
    print("download_dir:", cfg.download_dir)
    print("output_dir  :", cfg.output_dir)
    print("Tip: 카톡창 클릭 -> F8 -> 파일 저장/다운로드")

    ctx = Context()
    stop_event = threading.Event()

    # 핫키 스레드 시작
    t = threading.Thread(target=hotkey_thread_fn, args=(ctx, stop_event), daemon=True)
    t.start()

    # watchdog 시작
    obs = Observer()
    obs.schedule(Handler(cfg, ctx), cfg.download_dir, recursive=False)
    obs.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        stop_event.set()
        obs.stop()
        obs.join()


if __name__ == "__main__":
    main()
