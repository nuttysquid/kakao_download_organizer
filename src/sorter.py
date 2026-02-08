import json
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
    # repo 기준 상대경로로 config.json을 읽습니다.
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    cfg_path = os.path.join(base_dir, "config", "config.json")

    with open(cfg_path, "r", encoding="utf-8") as f:
        raw = json.load(f)

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

    return Config(
        download_dir=download_dir,
        output_dir=output_dir,
        hotkey_context_ttl_seconds=int(raw.get("hotkey_context_ttl_seconds", 180)),
        context_file=os.path.expandvars(raw.get("context_file", r"%TEMP%\kakao_room_ctx.txt")),
        rename_template=raw.get("rename_template", "{ts}__{room}__{bucket}__{orig}"),
        buckets=raw.get("buckets", {}),
        ignore_ext=[e.lower() for e in raw.get("ignore_ext", [])],
        log_dir=raw.get("log_dir", "logs"),
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
    def __init__(self, cfg: Config):
        super().__init__()
        self.cfg = cfg

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

        room, ts = get_room_context(self.cfg)
        ts = ts[:12]  # YYYYMMDDHHMM (초 제거)

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
            shutil.move(src, dst)
            log_line(self.cfg, f"MOVED {src} -> {dst}")
        except Exception as ex:
            log_line(self.cfg, f"FAIL move: {src} ({ex})")


def main():
    cfg = load_config()
    ensure_dirs(cfg)

    print("=== Kakao Download Organizer ===")
    print("download_dir:", cfg.download_dir)
    print("output_dir  :", cfg.output_dir)
    print("context_file:", cfg.context_file)
    print("Tip: 카톡창 클릭 -> F8 -> 파일 저장/다운로드")

    log_line(cfg, "START sorter")

    obs = Observer()
    obs.schedule(Handler(cfg), cfg.download_dir, recursive=False)
    obs.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        obs.stop()
        obs.join()
        log_line(cfg, "STOP sorter")


if __name__ == "__main__":
    main()
