@echo off
REM ================================================================
REM  HWP Converter - HWPX/HWP 변환 + 배포용 HWP 생성 (단건/다건)
REM  자세한 사용법은 README.txt 참조
REM ================================================================
setlocal enableextensions
chcp 65001 > nul

set "SCRIPT_DIR=%~dp0"
set "JAR=%SCRIPT_DIR%hwpConverter.jar"
set "LIB=%SCRIPT_DIR%lib\*"
set "CP=%JAR%;%LIB%"

where java >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Java not found on PATH. Please install Java 8+ and retry.
    exit /b 1
)

if "%~1"=="" goto :usage

REM --- --dist 모드 인자 개수 사전 검사 --------------------------------
REM  단건: --dist <input> <output> <password> [--no-copy] [--no-print]
REM  다건: --dist <inputDir> <outputDir> <password> [--no-copy] [--no-print]
if /I "%~1"=="--dist" (
    if "%~2"=="" goto :dist_args_err
    if "%~3"=="" goto :dist_args_err
    if "%~4"=="" goto :dist_args_err
)

REM --- 실행 --------------------------------------------------------
REM  다건 파일 지정 모드는 인자가 9개를 초과할 수 있으므로 %* 를 사용.
REM  단건/폴더배치는 9개 이하이므로 %1~%9 로도 커버 가능.
java -Xmx4g -Dfile.encoding=UTF-8 -Dstdout.encoding=UTF-8 -Dstderr.encoding=UTF-8 -cp "%CP%" kr.n.nframe.HwpConverter %*
set EXITCODE=%ERRORLEVEL%
if %EXITCODE% EQU 2 (
    echo.
    echo [TIP] --dist 인자에 공백/특수문자가 있다면 반드시 큰따옴표로 감싸세요.
    echo        예: hwpConverter.bat --dist "C:\Path With Spaces\in.hwp" "C:\Out\out.hwp" "mypw" --no-print
)
endlocal & exit /b %EXITCODE%

:dist_args_err
echo.
echo [ERROR] --dist 모드는 최소 3개의 위치 인자가 필요합니다.
echo   단건: hwpConverter.bat --dist ^<input.hwp^>  ^<output.hwp^> ^<password^> [--no-copy] [--no-print]
echo   다건: hwpConverter.bat --dist ^<inputDir^>   ^<outputDir^> ^<password^> [--no-copy] [--no-print]
echo.
echo  주의 1. 경로/암호에 공백이 있으면 반드시 큰따옴표로 감싸야 하며
echo          큰따옴표 짝이 홀수개이면 cmd 가 인자를 합쳐버립니다.
echo  주의 2. 암호에 ^< ^> ^& ^| 가 들어가면 cmd 의 리다이렉션으로 오인식될 수 있습니다.
echo          큰따옴표로 감싸되 caret(^^) 로 한 번 더 이스케이프 하세요:
echo            "^^<script^^>"    (실제 암호는 ^<script^>)
echo.
exit /b 2

:usage
echo.
echo  Usage (단건):
echo    hwpConverter.bat ^<input.hwpx^> ^<output.hwp^>                HWPX -^> HWP
echo    hwpConverter.bat ^<input.hwp^>  ^<output.hwpx^>               HWP  -^> HWPX
echo    hwpConverter.bat --dist ^<input^>  ^<output.hwp^>  ^<password^> [--no-copy] [--no-print]
echo.
echo  Usage (다건 - 폴더 전체):
echo    hwpConverter.bat ^<inputDir^> ^<outputDir^> [--to-hwpx ^| --to-hwp]
echo    hwpConverter.bat --dist ^<inputDir^> ^<outputDir^> ^<password^> [--no-copy] [--no-print] [--out-hwpx ^| --out-hwp]
echo.
echo  Usage (다건 - 개별 파일 지정):
echo    hwpConverter.bat ^<file1^> ^<file2^> ... ^<outputDir^> --to-hwpx ^| --to-hwp
echo    hwpConverter.bat --dist ^<file1^> ^<file2^> ... ^<outputDir^> ^<password^> [--no-copy] [--no-print] [--out-hwpx ^| --out-hwp]
echo.
echo  See README.txt for details.
exit /b 0
