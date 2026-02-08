import os
import time
import threading
import queue
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
    Statistics,
    History,
)


# -----------------------------
# Windows: 글로벌 핫키 등록
# -----------------------------
user32 = ctypes.windll.user32

WM_HOTKEY = 0x0312
HOTKEY_ID = 1
MOD_NOMOD = 0x0000

# 핫키 매핑 (Function Keys)
HOTKEY_VK_MAP = {
    "F1": 0x70,
    "F2": 0x71,
    "F3": 0x72,
    "F4": 0x73,
    "F5": 0x74,
    "F6": 0x75,
    "F7": 0x76,
    "F8": 0x77,
    "F9": 0x78,
    "F10": 0x79,
    "F11": 0x7A,
    "F12": 0x7B,
}

def get_hotkey_vk(hotkey_name: str) -> int:
    """핫키 이름을 Virtual Key 코드로 변환"""
    key = hotkey_name.upper()
    if key in HOTKEY_VK_MAP:
        return HOTKEY_VK_MAP[key]
    # 기본값은 F8
    print(f"Warning: Unknown hotkey '{hotkey_name}', using F8")
    return HOTKEY_VK_MAP["F8"]


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

# -----------------------------
# 작은 팝업(툴팁처럼) 표시용
# -----------------------------
_popup_q: "queue.Queue[tuple[str, int, int, int]]" = queue.Queue()

class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]

def get_foreground_hwnd() -> int:
    return int(user32.GetForegroundWindow())

def get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    if not hwnd:
        return None
    rect = RECT()
    ok = user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(rect))
    if not ok:
        return None
    return int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)

def popup_worker():
    # tkinter는 표준 포함이라 별도 설치 없이 사용 가능
    import tkinter as tk

    root = tk.Tk()
    root.withdraw()

    def poll():
        try:
            while True:
                text, x, y, ms = _popup_q.get_nowait()
                win = tk.Toplevel(root)
                win.overrideredirect(True)
                win.attributes("-topmost", True)

                # 심플한 작은 박스
                label = tk.Label(win, text=text, bg="black", fg="white", padx=10, pady=5)
                label.pack()

                win.update_idletasks()
                w = win.winfo_width()
                h = win.winfo_height()

                # 화면 밖으로 나가지 않게 약간 보정
                px = max(0, x - (w // 2))
                py = max(0, y)
                win.geometry(f"{w}x{h}+{px}+{py}")

                win.after(ms, win.destroy)
        except queue.Empty:
            pass

        root.after(80, poll)

    poll()
    root.mainloop()

def show_capture_popup(room: str):
    # 카톡 창 상단 근처에 0.9초 정도 팝업 표시
    hwnd = get_foreground_hwnd()
    r = get_window_rect(hwnd)
    if not r:
        return
    left, top, right, _ = r
    x = (left + right) // 2
    y = top + 12
    _popup_q.put((f"캡처됨: {room}", x, y, 900))


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


def hotkey_thread_fn(ctx: Context, stop_event: threading.Event, hotkey_name: str = "F8"):
    """
    글로벌 핫키를 등록하고 메시지 루프를 돌립니다.
    """
    vk_code = get_hotkey_vk(hotkey_name)

    # 기존 등록이 남아있을 수 있으니 해제 시도
    try:
        user32.UnregisterHotKey(None, HOTKEY_ID)
    except Exception:
        pass

    if not user32.RegisterHotKey(None, HOTKEY_ID, MOD_NOMOD, vk_code):
        # 등록 실패 (다른 프로그램이 점유했거나 권한/환경 문제)
        print(f"ERROR: {hotkey_name} 핫키 등록 실패. 다른 프로그램에서 {hotkey_name}을 사용 중일 수 있어요.")
        print("      (config.json에서 다른 핫키로 변경할 수 있습니다: F1~F12)")
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
            show_capture_popup(room)


    user32.UnregisterHotKey(None, HOTKEY_ID)


# -----------------------------
# 다운로드 폴더 감시 & 정리
# -----------------------------
class Handler(FileSystemEventHandler):
    def __init__(self, cfg, ctx: Context, stats=None, history=None):
        super().__init__()
        self.cfg = cfg
        self.ctx = ctx
        self.stats = stats
        self.history = history

    def _process_file(self, src: str):
        """공통 파일 처리 로직"""
        if not os.path.exists(src):
            return

        name = os.path.basename(src)
        _, ext = os.path.splitext(name)
        ext_l = ext.lower()

        # 무시할 확장자 체크
        if ext_l in self.cfg.ignore_ext:
            return

        # 제외할 확장자 체크
        if ext_l in self.cfg.exclude_extensions:
            log_line(self.cfg, f"SKIP excluded extension: {src}")
            return

        if not wait_until_ready(src):
            log_line(self.cfg, f"SKIP not-ready: {src}")
            return

        room, ts = self.ctx.get(self.cfg.hotkey_context_ttl_seconds)

        # 파일명에는 "분까지만" 쓰기 (YYYYMMDDHHMM)
        ts = ts[:12]

        room = safe_name(room)

        # 제외할 채팅방 체크
        if room in self.cfg.exclude_rooms:
            log_line(self.cfg, f"SKIP excluded room: {room}")
            return

        bucket = bucket_for_ext(self.cfg, ext_l)
        orig_safe = safe_name(name)

        new_name = self.cfg.rename_template.format(
            ts=ts, room=room, bucket=bucket, orig=orig_safe
        )
        new_name = safe_name(new_name)

        dst_dir = os.path.join(self.cfg.output_dir, room, bucket)
        os.makedirs(dst_dir, exist_ok=True)

        dst = os.path.join(dst_dir, new_name)

        # 중복 파일 처리
        if os.path.exists(dst):
            if self.cfg.duplicate_handling == "skip":
                log_line(self.cfg, f"SKIP duplicate: {dst}")
                return
            elif self.cfg.duplicate_handling == "overwrite":
                try:
                    os.remove(dst)
                    log_line(self.cfg, f"OVERWRITE: {dst}")
                except Exception as e:
                    log_line(self.cfg, f"FAIL overwrite: {dst} ({e})")
                    return
            else:  # rename (기본값)
                base, e2 = os.path.splitext(dst)
                i = 1
                while os.path.exists(dst):
                    dst = f"{base}({i}){e2}"
                    i += 1

        try:
            # 파일 크기 얻기 (통계용)
            file_size = 0
            try:
                file_size = os.path.getsize(src)
            except Exception:
                pass

            import shutil
            shutil.move(src, dst)
            log_line(self.cfg, f"MOVED {src} -> {dst}")

            # 통계 기록
            if self.stats:
                self.stats.record_file(room, bucket, file_size)

            # 히스토리 기록
            if self.history:
                self.history.record_move(src, dst, room, bucket)

        except Exception as ex:
            log_line(self.cfg, f"FAIL move: {src} ({ex})")

    def on_created(self, event):
        if event.is_directory:
            return
        self._process_file(event.src_path)

    def on_modified(self, event):
        if event.is_directory:
            return
        self._process_file(event.src_path)

    def on_moved(self, event):
        if event.is_directory:
            return
        # 파일이 이동되어 들어온 경우 (dest_path가 감시 폴더 내)
        self._process_file(event.dest_path)


# -----------------------------
# 자동 실행 & 바로가기 관리
# -----------------------------
def get_exe_path() -> str:
    """현재 실행 파일의 절대 경로를 반환"""
    import sys
    if getattr(sys, 'frozen', False):
        # PyInstaller로 빌드된 EXE
        return sys.executable
    else:
        # Python 스크립트로 실행 중
        return os.path.abspath(__file__)


def enable_autorun() -> bool:
    """Windows 시작 시 자동 실행 활성화"""
    try:
        import winreg
        exe_path = get_exe_path()
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "KakaoDownloadOrganizer", 0, winreg.REG_SZ, exe_path)

        print("✅ 자동 실행이 활성화되었습니다!")
        print(f"   경로: {exe_path}")
        return True
    except Exception as e:
        print(f"❌ 자동 실행 활성화 실패: {e}")
        return False


def disable_autorun() -> bool:
    """Windows 시작 시 자동 실행 비활성화"""
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            try:
                winreg.DeleteValue(key, "KakaoDownloadOrganizer")
                print("✅ 자동 실행이 비활성화되었습니다!")
                return True
            except FileNotFoundError:
                print("ℹ️  자동 실행이 설정되어 있지 않습니다.")
                return True
    except Exception as e:
        print(f"❌ 자동 실행 비활성화 실패: {e}")
        return False


def is_autorun_enabled() -> bool:
    """자동 실행 상태 확인"""
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
            try:
                value, _ = winreg.QueryValueEx(key, "KakaoDownloadOrganizer")
                return True
            except FileNotFoundError:
                return False
    except Exception:
        return False


def create_desktop_shortcut() -> bool:
    """바탕화면에 바로가기 생성"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")

        desktop = shell.SpecialFolders("Desktop")
        shortcut_path = os.path.join(desktop, "Kakao Download Organizer.lnk")

        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = get_exe_path()
        shortcut.WorkingDirectory = os.path.dirname(get_exe_path())
        shortcut.IconLocation = get_exe_path()
        shortcut.Description = "카카오톡 다운로드 파일 자동 정리"
        shortcut.save()

        print("✅ 바탕화면에 바로가기가 생성되었습니다!")
        print(f"   위치: {shortcut_path}")
        return True
    except ImportError:
        print("❌ pywin32가 설치되지 않았습니다.")
        print("   설치: pip install pywin32")
        return False
    except Exception as e:
        print(f"❌ 바로가기 생성 실패: {e}")
        return False


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Kakao Download Organizer")
    parser.add_argument("--autorun-enable", action="store_true", help="Windows 시작 시 자동 실행 활성화")
    parser.add_argument("--autorun-disable", action="store_true", help="Windows 시작 시 자동 실행 비활성화")
    parser.add_argument("--autorun-status", action="store_true", help="자동 실행 상태 확인")
    parser.add_argument("--create-shortcut", action="store_true", help="바탕화면 바로가기 생성")

    args = parser.parse_args()

    # 자동 실행 관련 명령어 처리
    if args.autorun_enable:
        enable_autorun()
        return

    if args.autorun_disable:
        disable_autorun()
        return

    if args.autorun_status:
        if is_autorun_enabled():
            print("✅ 자동 실행이 활성화되어 있습니다.")
        else:
            print("❌ 자동 실행이 비활성화되어 있습니다.")
        return

    if args.create_shortcut:
        create_desktop_shortcut()
        return

    # 일반 실행
    cfg = load_config()
    ensure_dirs(cfg)

    print("=== Kakao Download Organizer (single app) ===")
    print("download_dir:", cfg.download_dir)
    print("output_dir  :", cfg.output_dir)
    print(f"hotkey      : {cfg.hotkey}")
    print(f"Tip: 카톡창 클릭 -> {cfg.hotkey} -> 파일 저장/다운로드")

    ctx = Context()
    stop_event = threading.Event()

    # 통계 및 히스토리 초기화
    stats = Statistics(cfg) if cfg.enable_statistics else None
    history = History(cfg) if cfg.enable_history else None

    # 핫키 스레드 시작
    threading.Thread(target=popup_worker, daemon=True).start()
    t = threading.Thread(target=hotkey_thread_fn, args=(ctx, stop_event, cfg.hotkey), daemon=True)
    t.start()

    # watchdog 시작
    obs = Observer()
    obs.schedule(Handler(cfg, ctx, stats, history), cfg.download_dir, recursive=False)
    obs.start()

    try:
        # 주기적으로 통계 출력
        last_stats_print = time.time()
        while True:
            time.sleep(5)

            # 1시간마다 통계 출력
            if stats and time.time() - last_stats_print > 3600:
                print(stats.get_today_summary())
                last_stats_print = time.time()

    except KeyboardInterrupt:
        pass
    finally:
        stop_event.set()
        obs.stop()
        obs.join()

        # 종료 시 최종 통계 출력
        if stats:
            print(stats.get_today_summary())


if __name__ == "__main__":
    main()
