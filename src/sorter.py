import json
import sys
import os
import ctypes
from ctypes import wintypes

import time
import shutil
from dataclasses import dataclass
from typing import Dict, List, Tuple

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


def safe_name(s: str, max_len: int = 140) -> str:
    bad = '\\/:*?"<>|'
    for c in bad:
        s = s.replace(c, "_")
    s = s.strip()
    return s[:max_len] if len(s) > max_len else s


@dataclass
class Config:
    download_dir: str
    output_dir: str
    hotkey_context_ttl_seconds: int
    context_file: str
    rename_template: str
    buckets: Dict[str, List[str]]
    ignore_ext: List[str]
    log_dir: str
    # 새로운 설정들
    hotkey: str = "F8"
    exclude_rooms: List[str] = None
    exclude_extensions: List[str] = None
    duplicate_handling: str = "rename"  # rename, skip, overwrite
    enable_statistics: bool = True
    enable_history: bool = True

    def __post_init__(self):
        if self.exclude_rooms is None:
            self.exclude_rooms = []
        if self.exclude_extensions is None:
            self.exclude_extensions = []

class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", ctypes.c_ubyte * 8),
    ]


FOLDERID_Documents = GUID(
    0xFDD39AD0, 0x238F, 0x46AF, (ctypes.c_ubyte * 8)(0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7)
)
FOLDERID_Downloads = GUID(
    0x374DE290, 0x123F, 0x4565, (ctypes.c_ubyte * 8)(0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B)
)


def get_known_folder_path(folder_id: GUID) -> str:
    """
    Windows가 실제로 사용하는 Known Folder 경로를 반환합니다.
    (문서/다운로드가 OneDrive로 리디렉션된 경우도 반영)
    """
    # 비윈도우 환경 대비(안전장치)
    if os.name != "nt":
        return os.path.join(os.path.expandvars("%USERPROFILE%"), "Documents")

    SHGetKnownFolderPath = ctypes.windll.shell32.SHGetKnownFolderPath
    SHGetKnownFolderPath.argtypes = [ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)]
    SHGetKnownFolderPath.restype = wintypes.HRESULT

    p_path = ctypes.c_wchar_p()
    hr = SHGetKnownFolderPath(ctypes.byref(folder_id), 0, None, ctypes.byref(p_path))
    if hr != 0 or not p_path.value:
        # 실패 시 fallback
        return os.path.join(os.path.expandvars("%USERPROFILE%"), "Documents")

    path = p_path.value
    ctypes.windll.ole32.CoTaskMemFree(p_path)
    return path


def autodetect_kakaotalk_download_dir() -> str:
    """
    카카오톡 PC에서 흔히 쓰이는 기본 다운로드 폴더 후보를 찾아서
    '존재하는' 경로를 우선 반환합니다.
    """
    docs = get_known_folder_path(FOLDERID_Documents)
    dls = get_known_folder_path(FOLDERID_Downloads)

    candidates = [
        os.path.join(docs, "카카오톡 받은 파일"),
        os.path.join(dls, "KakaoTalk Downloads"),
        os.path.join(dls, "KakaoTalk"),
        os.path.join(dls, "KakaoTalk", "Download"),
        os.path.join(docs, "KakaoTalk Downloads"),
    ]

    for p in candidates:
        if os.path.isdir(p):
            return p

    # 아무 후보도 없으면: 문서 아래 기본값으로(프로그램이 폴더 생성 가능)
    return os.path.join(docs, "카카오톡 받은 파일")


def autodetect_output_dir() -> str:
    # 정리 결과는 "문서\KakaoSorted"로(OneDrive 포함 자동)
    docs = get_known_folder_path(FOLDERID_Documents)
    return os.path.join(docs, "KakaoSorted")


def load_config() -> Config:
    """
    - 개발(소스) 실행: repo/config/config.json 사용
    - EXE(배포) 실행: exe 옆의 config/config.json 또는 exe 옆 config.json 사용
    - 설정 파일이 없으면: 기본값(AUTO)으로 동작
    """
    # 1) 설정 파일 후보 경로들(우선순위)
    candidates: list[str] = []

    # EXE로 실행 중이면(= PyInstaller frozen)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        candidates += [
            os.path.join(exe_dir, "config", "config.json"),
            os.path.join(exe_dir, "config.json"),
        ]

    # 소스 실행일 때(또는 fallback)
    repo_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates += [
        os.path.join(repo_dir, "config", "config.json"),
    ]

    # 2) 기본 설정(설정 파일이 없을 때 사용)
    default_raw = {
        "download_dir": "AUTO",
        "output_dir": "AUTO",
        "hotkey_context_ttl_seconds": 180,
        "context_file": r"%TEMP%\kakao_room_ctx.txt",
        "rename_template": "{ts}__{room}__{bucket}__{orig}",
        "buckets": {
            "이미지": [".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"],
            "문서": [".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".hwp", ".hwpx", ".txt", ".rtf", ".csv"],
            "압축": [".zip", ".7z", ".rar"],
            "오디오": [".mp3", ".wav", ".m4a", ".aac", ".flac"],
            "비디오": [".mp4", ".mov", ".mkv", ".avi", ".webm"],
        },
        "ignore_ext": [".crdownload", ".tmp", ".part"],
        "log_dir": "logs",
    }

    # 3) 실제 설정 로드(있으면 읽고, 없으면 기본값)
    raw = None
    for p in candidates:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                raw = json.load(f)
            break

    if raw is None:
        raw = default_raw

    # 4) download/output 처리 (AUTO면 자동탐지, 아니면 환경변수 확장)
    download_raw = str(raw.get("download_dir", "AUTO")).strip()
    output_raw = str(raw.get("output_dir", "AUTO")).strip()

    download_dir = (
        autodetect_kakaotalk_download_dir()
        if download_raw.upper() == "AUTO"
        else os.path.expandvars(download_raw)
    )
    output_dir = (
        autodetect_output_dir()
        if output_raw.upper() == "AUTO"
        else os.path.expandvars(output_raw)
    )

    # 5) log_dir 처리 (EXE일 때는 exe 옆에 두는 게 덜 헷갈림)
    log_dir = os.path.expandvars(str(raw.get("log_dir", "logs")).strip())
    if getattr(sys, "frozen", False) and not os.path.isabs(log_dir):
        log_dir = os.path.join(os.path.dirname(sys.executable), log_dir)

    return Config(
        download_dir=download_dir,
        output_dir=output_dir,
        hotkey_context_ttl_seconds=int(raw.get("hotkey_context_ttl_seconds", 180)),
        context_file=os.path.expandvars(raw.get("context_file", r"%TEMP%\kakao_room_ctx.txt")),
        rename_template=raw.get("rename_template", "{ts}__{room}__{bucket}__{orig}"),
        buckets=raw.get("buckets", {}),
        ignore_ext=[e.lower() for e in raw.get("ignore_ext", [])],
        log_dir=log_dir,
        # 새로운 설정들
        hotkey=str(raw.get("hotkey", "F8")).upper(),
        exclude_rooms=raw.get("exclude_rooms", []),
        exclude_extensions=[e.lower() for e in raw.get("exclude_extensions", [])],
        duplicate_handling=raw.get("duplicate_handling", "rename"),
        enable_statistics=bool(raw.get("enable_statistics", True)),
        enable_history=bool(raw.get("enable_history", True)),
    )


def ensure_dirs(cfg: Config) -> None:
    os.makedirs(cfg.download_dir, exist_ok=True)
    os.makedirs(cfg.output_dir, exist_ok=True)
    os.makedirs(cfg.log_dir, exist_ok=True)


def log_line(cfg: Config, msg: str) -> None:
    day = time.strftime("%Y-%m-%d")
    path = os.path.join(cfg.log_dir, f"sorter_{day}.log")
    line = f"[{time.strftime('%H:%M:%S')}] {msg}\n"
    with open(path, "a", encoding="utf-8") as f:
        f.write(line)


# 통계 추적 클래스
class Statistics:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.stats_file = os.path.join(cfg.log_dir, "statistics.json")
        self.today = time.strftime("%Y-%m-%d")
        self.stats = self._load_stats()

    def _load_stats(self) -> dict:
        if os.path.exists(self.stats_file):
            try:
                with open(self.stats_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_stats(self):
        try:
            with open(self.stats_file, "w", encoding="utf-8") as f:
                json.dump(self.stats, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log_line(self.cfg, f"Failed to save stats: {e}")

    def record_file(self, room: str, bucket: str, file_size: int):
        if not self.cfg.enable_statistics:
            return

        if self.today not in self.stats:
            self.stats[self.today] = {}

        day_stats = self.stats[self.today]

        # 전체 통계
        if "total" not in day_stats:
            day_stats["total"] = {"count": 0, "size": 0}
        day_stats["total"]["count"] += 1
        day_stats["total"]["size"] += file_size

        # 버킷별 통계
        if "by_bucket" not in day_stats:
            day_stats["by_bucket"] = {}
        if bucket not in day_stats["by_bucket"]:
            day_stats["by_bucket"][bucket] = {"count": 0, "size": 0}
        day_stats["by_bucket"][bucket]["count"] += 1
        day_stats["by_bucket"][bucket]["size"] += file_size

        # 채팅방별 통계
        if "by_room" not in day_stats:
            day_stats["by_room"] = {}
        if room not in day_stats["by_room"]:
            day_stats["by_room"][room] = {"count": 0, "size": 0}
        day_stats["by_room"][room]["count"] += 1
        day_stats["by_room"][room]["size"] += file_size

        self._save_stats()

    def get_today_summary(self) -> str:
        if self.today not in self.stats:
            return "오늘 정리된 파일: 0개"

        day_stats = self.stats[self.today]
        total = day_stats.get("total", {"count": 0, "size": 0})
        count = total["count"]
        size_mb = total["size"] / (1024 * 1024)

        lines = [f"\n📊 오늘 정리된 파일: {count}개 ({size_mb:.1f}MB)"]

        by_bucket = day_stats.get("by_bucket", {})
        if by_bucket:
            lines.append("  종류별:")
            for bucket, data in by_bucket.items():
                bcount = data["count"]
                bsize = data["size"] / (1024 * 1024)
                lines.append(f"    {bucket}: {bcount}개 ({bsize:.1f}MB)")

        return "\n".join(lines)


# 히스토리 추적 클래스
class History:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.history_file = os.path.join(cfg.log_dir, "history.json")
        self.today = time.strftime("%Y-%m-%d")
        self.history = self._load_history()

    def _load_history(self) -> dict:
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_history(self):
        try:
            # 최근 30일만 보관
            cutoff = time.time() - (30 * 24 * 60 * 60)
            cutoff_date = time.strftime("%Y-%m-%d", time.localtime(cutoff))

            filtered = {k: v for k, v in self.history.items() if k >= cutoff_date}

            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(filtered, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log_line(self.cfg, f"Failed to save history: {e}")

    def record_move(self, src: str, dst: str, room: str, bucket: str):
        if not self.cfg.enable_history:
            return

        if self.today not in self.history:
            self.history[self.today] = []

        self.history[self.today].append({
            "time": time.strftime("%H:%M:%S"),
            "src": src,
            "dst": dst,
            "room": room,
            "bucket": bucket
        })

        self._save_history()


def get_room_context(cfg: Config) -> Tuple[str, str]:
    try:
        with open(cfg.context_file, "r", encoding="utf-8") as f:
            raw = f.read().strip()
        room, ts = raw.split("|", 1)

        t_struct = time.strptime(ts, "%Y%m%d%H%M%S")
        age = time.time() - time.mktime(t_struct)

        if age <= cfg.hotkey_context_ttl_seconds:
            return room, ts
    except Exception:
        pass

    return "미분류", time.strftime("%Y%m%d%H%M%S")


def bucket_for_ext(cfg: Config, ext: str) -> str:
    e = ext.lower()
    for bucket, exts in cfg.buckets.items():
        if e in [x.lower() for x in exts]:
            return bucket
    return "기타"


def wait_until_ready(path: str, timeout_sec: int = 20) -> bool:
    start = time.time()
    last_size = -1
    stable = 0

    while time.time() - start < timeout_sec:
        try:
            size = os.path.getsize(path)
            if size == last_size and size > 0:
                stable += 1
            else:
                stable = 0
            last_size = size

            if stable >= 3:
                with open(path, "rb"):
                    return True
        except Exception:
            pass
        time.sleep(0.2)

    return False


class Handler(FileSystemEventHandler):
    def __init__(self, cfg: Config, stats: Statistics = None, history: History = None):
        super().__init__()
        self.cfg = cfg
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

        room, ts = get_room_context(self.cfg)
        ts = ts[:12]  # YYYYMMDDHHMM (초 제거)

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


def main():
    cfg = load_config()
    ensure_dirs(cfg)

    print("=== Kakao Download Organizer ===")
    print("download_dir:", cfg.download_dir)
    print("output_dir  :", cfg.output_dir)
    print("context_file:", cfg.context_file)
    print(f"hotkey      : {cfg.hotkey}")
    print("Tip: 카톡창 클릭 -> F8 -> 파일 저장/다운로드")

    log_line(cfg, "START sorter")

    # 통계 및 히스토리 초기화
    stats = Statistics(cfg) if cfg.enable_statistics else None
    history = History(cfg) if cfg.enable_history else None

    obs = Observer()
    obs.schedule(Handler(cfg, stats, history), cfg.download_dir, recursive=False)
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
        obs.stop()
        obs.join()

        # 종료 시 최종 통계 출력
        if stats:
            print(stats.get_today_summary())

        log_line(cfg, "STOP sorter")


if __name__ == "__main__":
    main()
