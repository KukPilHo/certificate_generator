@echo off
REM 더블클릭하면 수료증을 일괄 생성합니다. (Windows)
REM 이 파일은 프로젝트 폴더 안에 그대로 두세요.
chcp 65001 >nul
setlocal

REM 1) 이 배치 파일이 있는 폴더로 이동
cd /d "%~dp0"
echo 작업 폴더: %cd%
echo ----------------------------------------

REM 2) Python 인터프리터 선택 (venv 우선)
set "PY=python"
if exist "venv\Scripts\python.exe" set "PY=venv\Scripts\python.exe"
echo 사용 Python: %PY%

REM 3) 필수 패키지 확인 후 없으면 설치
"%PY%" -c "import openpyxl, pptx" 2>nul
if errorlevel 1 (
  echo 필수 패키지를 설치합니다...
  "%PY%" -m pip install -r requirements.txt
)

REM 4) 입력 파일 자동 탐색 (폴더 안 첫 .xlsx / .pptx)
set "EXCEL="
set "TEMPLATE="
for %%f in ("certificate excel\*.xlsx") do set "EXCEL=%%f"
for %%f in ("certificate ppt\*.pptx") do set "TEMPLATE=%%f"
if not defined EXCEL (
  echo [오류] 'certificate excel' 폴더에서 .xlsx 파일을 찾지 못했습니다.
  pause & exit /b 1
)
if not defined TEMPLATE (
  echo [오류] 'certificate ppt' 폴더에서 .pptx 파일을 찾지 못했습니다.
  pause & exit /b 1
)
echo 엑셀:   %EXCEL%
echo 템플릿: %TEMPLATE%

REM 5) 수료증 생성 실행
echo ----------------------------------------
"%PY%" generate_certificates.py --excel "%EXCEL%" --template "%TEMPLATE%"
set STATUS=%errorlevel%
echo ----------------------------------------
if "%STATUS%"=="0" (
  echo [완료] 결과: 'output ppt' / 'output pdf' 폴더를 확인하세요.
) else (
  echo [오류] 코드 %STATUS%. 위 메시지를 확인하세요.
)

echo.
pause
endlocal
