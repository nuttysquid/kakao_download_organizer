@echo off
setlocal
cd /d "%~dp0\.."

REM 1) AHK 실행 (F8 캡처)
start "" "%cd%\ahk\room_capture.ahk"

REM 2) Python sorter 실행 (다운로드 폴더 감시)
if exist "%cd%\.venv\Scripts\python.exe" (
  "%cd%\.venv\Scripts\python.exe" "%cd%\src\sorter.py"
) else (
  python "%cd%\src\sorter.py"
)

endlocal
