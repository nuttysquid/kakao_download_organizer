# Kakao Download Organizer (Windows)

PC 카카오톡에서 다운로드/저장되는 파일을 **채팅방별 + 파일종류별**로 자동 정리합니다.

- 기본 다운로드 폴더를 자동으로 찾습니다(OneDrive 문서 포함).
- 사용 방법: **채팅방 클릭 → F8 → 파일 저장/다운로드**

### 다운로드(아래 두가지 방법 중에 선택하여 다운로드. 초보자는 1번 추천)
### ✅ 1.Python/AutoHotkey 설치 없이 쓰기(Windows EXE)
1) 이 저장소의 **Releases**로 이동
2) 최신 버전에서 **`KakaoDownloadOrganizer_v0.1.1_windows.zip`** 다운로드
3) 압축 해제 후 `KakaoDownloadOrganizer.exe` 실행  
4) 카톡 채팅방 클릭 → **F8** → 파일 저장/다운로드

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

## 자동 실행(선택: Windows 시작 시 자동 실행)

항상 자동 정리를 쓰고 싶다면, Windows 시작 시 자동 실행되게 설정할 수 있습니다.

1. `Win + R` → `shell:startup` 입력 → Enter
2. 열리는 “시작프로그램” 폴더에 `scripts\start.bat` **바로가기**를 넣기
3. (추천) 바로가기 우클릭 → 속성 → **실행: 최소화된 창** 으로 변경

이 설정을 하면 PC를 켤 때 자동으로 실행됩니다.
(현재 버전은 콘솔(검은 창)이 최소화 상태로 실행됩니다.)

## 설정(선택)

- `config/config.json`에서 `"download_dir"`/`"output_dir"`를 직접 지정할 수 있습니다.
- 기본값은 `"AUTO"`이며, OneDrive 환경도 자동 대응합니다.

## 주의

- 이 도구는 카카오톡 “채팅 내용”을 읽지 않습니다.
- 네트워크로 데이터를 전송하지 않습니다. (`docs/PRIVACY.md` 참고)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install watchdog
