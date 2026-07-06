@echo off
REM ================================================================
REM  jar 매니페스트의 Class-Path 헤더를 제거합니다 (소스 재컴파일 불필요)
REM
REM  사용법:  strip-classpath.bat ^<target.jar^>
REM    예) strip-classpath.bat hwpConverter.jar
REM ================================================================
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo Usage: strip-classpath.bat ^<jarfile^>
    exit /b 1
)
set "JAR=%~f1"
if not exist "%JAR%" (
    echo [ERROR] not found: %JAR%
    exit /b 2
)

where jar >nul 2>nul
if errorlevel 1 (
    echo [ERROR] jar tool not found on PATH. Install JDK and retry.
    exit /b 3
)

REM 백업
copy /y "%JAR%" "%JAR%.bak" >nul

REM 임시 디렉토리
set "WORK=%TEMP%\stripcp_%RANDOM%"
mkdir "%WORK%"
pushd "%WORK%"

REM 1) 매니페스트 추출
jar xf "%JAR%" META-INF/MANIFEST.MF
if errorlevel 1 goto :err

REM 2) Class-Path 와 continuation 줄 제거 (PowerShell 사용)
powershell -NoProfile -Command "$lines = Get-Content -Encoding UTF8 'META-INF/MANIFEST.MF'; $skip = $false; $out = foreach ($l in $lines) { if ($l -match '^Class-Path:') { $skip = $true; continue } elseif ($skip -and ($l -match '^[ \t]')) { continue } else { $skip = $false; $l } }; $out | Set-Content -Encoding UTF8 'META-INF/MANIFEST.MF'"
if errorlevel 1 goto :err

REM 3) jar 에 매니페스트 다시 넣기
jar uf "%JAR%" META-INF/MANIFEST.MF
if errorlevel 1 goto :err

popd
rmdir /s /q "%WORK%"

echo.
echo [OK] Class-Path stripped from: %JAR%
echo      Backup: %JAR%.bak
echo.
echo --- updated manifest ---
jar xf "%JAR%" META-INF/MANIFEST.MF
type META-INF/MANIFEST.MF
rmdir /s /q META-INF
exit /b 0

:err
popd 2>nul
echo [ERROR] failed. Backup retained at: %JAR%.bak
exit /b 9
