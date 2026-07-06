@echo off
REM ============================================================
REM  hwpConverter API 테스트 컴파일 + 실행
REM ============================================================
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "ROOT=%SCRIPT_DIR%.."
set "JAR=%ROOT%\hwpConverter.jar"
set "LIB=%ROOT%\lib\*"

if not exist "%JAR%" (
    echo [ERROR] hwpConverter.jar not found: %JAR%
    exit /b 1
)

REM 1) 컴파일
if not exist "%SCRIPT_DIR%build" mkdir "%SCRIPT_DIR%build"
javac -encoding UTF-8 -d "%SCRIPT_DIR%build" -cp "%JAR%;%LIB%" "%SCRIPT_DIR%HwpConverterApiTest.java"
if errorlevel 1 (
    echo [ERROR] compile failed.
    exit /b 2
)

REM 2) 실행 — 작업 디렉토리를 apitest\ 로 두고 ../hwp/pages 의 입력을 사용
pushd "%SCRIPT_DIR%"
java -Dfile.encoding=UTF-8 -Dstdout.encoding=UTF-8 -cp "build;%JAR%;%LIB%" HwpConverterApiTest
set EXITCODE=%ERRORLEVEL%
popd

endlocal & exit /b %EXITCODE%
