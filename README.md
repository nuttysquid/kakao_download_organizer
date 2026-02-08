# Kakao Download Organizer (Windows)

PC 카카오톡에서 다운로드/저장되는 파일을 **채팅방별 + 파일종류별**로 자동 정리합니다.

- 기본 다운로드 폴더를 자동으로 찾습니다(OneDrive 문서 포함).
- 사용 방법: **채팅방 클릭 → F8 → 파일 저장/다운로드**

### 다운로드(아래 두가지 방법 중에 선택하여 다운로드. 초보자는 1번 추천)
### ✅ 1.Python/AutoHotkey 설치 없이 쓰기(Windows EXE)
1) 이 저장소의 **Releases**로 이동
2) 최신 버전에서 **`KakaoDownloadOrganizer_v0.3.0_windows.zip`** 다운로드
3) 압축 해제 후 `KakaoDownloadOrganizer.exe` 실행
4) 카톡 채팅방 클릭 → **F8** → 파일 저장/다운로드

**✨ v0.3.0 신기능:**
- 🚀 **Windows 시작 시 자동 실행** (명령줄 인자로 쉽게 설정)
- 🔗 **바탕화면 바로가기 생성**

**v0.2.x 기능:**
- 핫키 커스터마이징 (F1~F12)
- 실시간 통계 표시
- 파일 이동 히스토리 (30일)
- 채팅방/확장자 제외 기능
- 중복 파일 처리 옵션

※ EXE 전용 안내서는 `README_EXE.md`를 참고하세요.



### ✅2.개발/커스터마이징(소스 버전: Python + AutoHotkey)
아래 “설치/실행” 안내(README 본문)를 따라 설치해서 사용하세요.

### 1) 준비물
- Windows
- Python 3.x 설치
- AutoHotkey v2 설치

### 2) 설치
1. 이 저장소를 다운로드(zip)하거나 Git으로 clone 합니다.
2. 폴더에서 터미널을 열고 아래를 실행합니다.

### 3) 실행

실행 명령:
`.\scripts\start.bat`

### 4) 사용 방법

1. PC 카카오톡에서 정리하고 싶은 채팅방을 클릭(활성화)
2. **F8** 누르기(채팅방 캡처)
3. 파일을 저장/다운로드  
   → 자동으로 정리 폴더로 이동됩니다.

정리 폴더(기본):
- `문서\KakaoSorted\<채팅방>\<종류>\...`

## 편의 기능 (v0.3.0+)

### 🚀 Windows 시작 시 자동 실행

**EXE 버전 (간편):**
```powershell
# 자동 실행 활성화
KakaoDownloadOrganizer.exe --autorun-enable

# 자동 실행 비활성화
KakaoDownloadOrganizer.exe --autorun-disable

# 상태 확인
KakaoDownloadOrganizer.exe --autorun-status
```

**소스 버전:**
```powershell
python src/app.py --autorun-enable
```

### 🔗 바탕화면 바로가기 만들기

```powershell
# EXE 버전
KakaoDownloadOrganizer.exe --create-shortcut

# 소스 버전 (pywin32 필요)
pip install pywin32
python src/app.py --create-shortcut
```

### 수동 설정 (구버전 호환)

1. `Win + R` → `shell:startup` 입력 → Enter
2. 열리는 "시작프로그램" 폴더에 `KakaoDownloadOrganizer.exe` 바로가기 넣기
3. (추천) 바로가기 우클릭 → 속성 → **실행: 최소화된 창**

## 설정(선택)

- `config/config.json`에서 다양한 설정을 변경할 수 있습니다.
- 기본값은 `"AUTO"`이며, OneDrive 환경도 자동 대응합니다.

### 주요 설정 옵션

```json
{
  "download_dir": "AUTO",
  "output_dir": "AUTO",

  "hotkey": "F8",

  "exclude_rooms": ["광고방", "스팸방"],
  "exclude_extensions": [".exe", ".msi", ".bat"],

  "duplicate_handling": "rename",

  "enable_statistics": true,
  "enable_history": true
}
```

**설정 설명:**
- `hotkey`: 채팅방 캡처 단축키 (F1~F12)
- `exclude_rooms`: 정리하지 않을 채팅방 이름
- `exclude_extensions`: 정리하지 않을 파일 확장자
- `duplicate_handling`: 중복 파일 처리 방법
  - `"rename"`: 파일명에 (1), (2) 추가 (기본값)
  - `"skip"`: 건너뛰기
  - `"overwrite"`: 덮어쓰기
- `enable_statistics`: 통계 기능 켜기/끄기
- `enable_history`: 히스토리 기능 켜기/끄기

## 주의

- 이 도구는 카카오톡 “채팅 내용”을 읽지 않습니다.
- 네트워크로 데이터를 전송하지 않습니다. (`docs/PRIVACY.md` 참고)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install watchdog
