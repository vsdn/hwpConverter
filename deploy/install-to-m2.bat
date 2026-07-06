@echo off
REM ============================================================
REM  hwpConverter + 의존 jar 11개를 로컬 Maven 저장소(~/.m2)에 설치
REM  사용 전 mvn 이 PATH 에 있어야 합니다.
REM ============================================================
setlocal
set "BASE=%~dp0.."

where mvn >nul 2>nul
if errorlevel 1 (
    echo [ERROR] mvn not found on PATH.
    exit /b 1
)

call mvn install:install-file -Dfile="%BASE%\hwpConverter.jar"        -DgroupId=kr.n.nframe   -DartifactId=hwpConverter   -Dversion=13.15  -Dpackaging=jar
call mvn install:install-file -Dfile="%BASE%\lib\hwplib-1.1.10.jar"   -DgroupId=kr.dogfoot    -DartifactId=hwplib         -Dversion=1.1.10 -Dpackaging=jar
call mvn install:install-file -Dfile="%BASE%\lib\hwpxlib-1.0.9.jar"   -DgroupId=kr.dogfoot    -DartifactId=hwpxlib        -Dversion=1.0.9  -Dpackaging=jar
call mvn install:install-file -Dfile="%BASE%\lib\hwp2hwpx-1.0.0.jar"  -DgroupId=kr.dogfoot    -DartifactId=hwp2hwpx       -Dversion=1.0.0  -Dpackaging=jar

REM 나머지 8개는 Maven Central 에 그대로 있으니 pom.xml 의존성으로 충분.
echo.
echo [OK] hwpConverter, hwplib, hwpxlib, hwp2hwpx 4개를 ~/.m2 에 설치했습니다.
echo      나머지 의존성(poi, log4j, commons-*, SparseBitSet)은 Maven Central
echo      에서 자동 해결되므로 pom.xml 에 dependency 만 추가하면 됩니다.
endlocal
