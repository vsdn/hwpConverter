WEB-INF/lib 배포용 가이드
==========================

다음 12개 jar 파일을 소비측 웹 애플리케이션의 WEB-INF/lib/ 에 그대로
복사하면 됩니다. 추가 설정 없이 서블릿 컨테이너의 클래스로더가 모두
인식합니다.

  1) hwpConverter.jar             (이 프로젝트 본체)
  2) lib/hwplib-1.1.10.jar
  3) lib/hwpxlib-1.0.9.jar
  4) lib/hwp2hwpx-1.0.0.jar
  5) lib/poi-5.2.5.jar
  6) lib/log4j-api-2.22.1.jar
  7) lib/log4j-core-2.22.1.jar
  8) lib/commons-codec-1.16.0.jar
  9) lib/commons-collections-3.2.2.jar
 10) lib/commons-io-2.15.1.jar
 11) lib/commons-math3-3.6.1.jar
 12) lib/SparseBitSet-1.3.jar

주의 사항
----------
 * hwpConverter.jar 의 META-INF/MANIFEST.MF 에는 Class-Path 헤더가 없습니다.
   (있으면 일부 컨테이너의 jar 스캐너가 상대 경로 lib/ 을 따라가다 실패합니다.)

 * 호스트 프로젝트가 이미 같은 라이브러리를 다른 버전으로 가지고 있으면
   클래스로더 충돌을 일으킬 수 있습니다. 특히 POI / log4j / commons-* 는
   다른 라이브러리도 자주 사용하므로 사전 확인 필요.

 * Servlet 환경의 임시 파일 경로는 java.io.tmpdir 시스템 속성 또는
   ServletContext getRealPath 로 확보하세요. 본 라이브러리는 변환 도중
   /tmp 와 같은 OS 기본 임시 경로에 임시 파일을 만들 수 있습니다.

 * Java 런타임은 8 이상이어야 합니다 (모든 jar 가 Java 8 바이트코드).

호출 예시
----------
    import kr.n.nframe.HwpConverter;

    HwpConverter conv = new HwpConverter();
    conv.convertHwpxToHwp("input.hwpx", "output.hwp");

    // 배포용(DRM)
    conv.makeHwpDist("input.hwp", "output.hwp",
                     "password",
                     /*noCopy*/ false, /*noPrint*/ false);
