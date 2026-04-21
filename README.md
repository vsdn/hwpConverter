# hwpConverter

github.com/neolord0님의 kr.dogfoot:hwplib 공유에 감사드립니다.

HWP · HWPX 양방향 변환기와 한글 배포용(DRM) 문서 변환기.

- **HWPX → HWP** : 직접 바이너리 구현 (자체 writer)
- **HWP → HWPX** : `kr.dogfoot:hwplib` / `hwpxlib` / `hwp2hwpx` 기반 + 후처리
- **HWP → 배포용 HWP** : AES-128 암호화 ViewText 스트림 생성 (DRM 플래그 · 복사/인쇄 방지)

CLI로도 지원하고 라이브러리로도 지원합니다.

문의 혹은 버그 리포트는 cowork@nsoftware.kr 로 전달바랍니다.

## 요구사항

- **Java 8 이상** (`javac --release 8` 타겟)
- Apache POI 5.2.5 (OLE2 컨테이너)
- 외부 JAR 은 `lib/` 폴더에 포함되어 있음

## 빌드

```bash
# 소스 목록 갱신 (필요 시)
find src -name "*.java" > build/sources.txt

# 컴파일 + 패키징
javac --release 8 -d build/classes -cp "lib/*" -encoding UTF-8 @build/sources.txt
jar cfm build/hwpConverter.jar build/MANIFEST.MF -C build/classes kr
```

Eclipse 사용자는 `.classpath` · `.project` 가 포함되어 바로 임포트 가능.

## 실행

### Windows
```cmd
hwpConverter.bat <input.hwpx> <output.hwp>
```

### Linux / macOS
```bash
chmod +x release/hwpConverter.sh   # 최초 1회
./release/hwpConverter.sh <input.hwpx> <output.hwp>
```

### 사용 예

| 작업 | 명령 |
|------|------|
| HWPX → HWP | `hwpConverter.bat <in.hwpx> <out.hwp>` |
| HWP → HWPX | `hwpConverter.bat <in.hwp> <out.hwpx>` |
| 배포용(DRM) | `hwpConverter.bat --dist <in.hwp> <out.hwp> <password> [--no-copy] [--no-print]` |
| 폴더 배치 | `hwpConverter.bat <inputDir> <outputDir> [--to-hwpx\|--to-hwp]` |
| 파일 지정 | `hwpConverter.bat <f1> <f2> ... <outputDir> --to-hwpx\|--to-hwp]` |
| 배포용 다건변경 (DRM) | `hwpConverter.bat --dist <f1> <f2> ... <outputDir> <password> [--no-copy] [--no-print] [--out-hwpx\--out-hwp]` |

경로·암호에 공백이나 셸 특수문자(`< > & |`) 가 있으면 반드시 따옴표로 감싸세요.


## 라이브러리로 사용하기

외부 Java 프로젝트에서 변환 메서드를 직접 호출할 수 있습니다.

### 1. 의존성 추가

**Maven / Gradle 없이 (release 배포본 사용)**

`hwpConverter-v1.0.zip` 압축을 풀어 `hwpConverter.jar` + `lib/*.jar` 를 classpath 에 추가.

```bash
# 컴파일
javac -cp "hwpConverter.jar;lib/*" MyApp.java

# 실행 (Windows: ';', Linux/macOS: ':')
java  -cp "hwpConverter.jar;lib/*;." MyApp
java  -cp "hwpConverter.jar:lib/*:." MyApp
```

**Maven (로컬 설치)**

```bash
mvn install:install-file -Dfile=hwpConverter.jar \
    -DgroupId=kr.n.nframe -DartifactId=hwpConverter \
    -Dversion=13.15 -Dpackaging=jar
```

```xml
<dependency>
    <groupId>kr.n.nframe</groupId>
    <artifactId>hwpConverter</artifactId>
    <version>13.15</version>
</dependency>

<!-- 외부 의존성 (lib/ 에 포함된 것들) -->
<dependency><groupId>org.apache.poi</groupId><artifactId>poi</artifactId><version>5.2.5</version></dependency>
<dependency><groupId>kr.dogfoot</groupId><artifactId>hwplib</artifactId><version>1.1.10</version></dependency>
<dependency><groupId>kr.dogfoot</groupId><artifactId>hwpxlib</artifactId><version>1.0.9</version></dependency>
<dependency><groupId>kr.dogfoot</groupId><artifactId>hwp2hwpx</artifactId><version>1.0.0</version></dependency>
```

**Gradle**

```groovy
dependencies {
    implementation files('libs/hwpConverter.jar')
    implementation fileTree(dir: 'libs/lib', include: ['*.jar'])
}
```

### 2. 기본 변환 API

```java
import kr.n.nframe.HwpConverter;

public class Example {
    public static void main(String[] args) throws Exception {
        HwpConverter converter = new HwpConverter();

        // HWPX → HWP (직접 바이너리 변환)
        converter.convertHwpxToHwp("input.hwpx", "output.hwp");

        // HWP → HWPX (hwp2hwpx 기반)
        converter.convertHwpToHwpx("input.hwp", "output.hwpx");

        // HWP → 배포용(DRM) HWP — AES-128 암호화 + 복사/인쇄 방지
        String password = "mySecret";
        boolean noCopy  = true;   // 복사 방지 활성화
        boolean noPrint = false;  // 인쇄 허용
        converter.makeHwpDist("input.hwp", "output_dist.hwp",
                              password, noCopy, noPrint);
    }
}
```

### 3. 배치 변환 API

```java
import kr.n.nframe.HwpConverter;
import java.io.File;
import java.util.Arrays;
import java.util.List;

HwpConverter c = new HwpConverter();

// 폴더 전체 변환 (재귀 아님, 해당 디렉터리의 .hwpx 만)
c.batchHwpxToHwp("in_dir", "out_dir");
c.batchHwpToHwpx("in_dir", "out_dir");

// 배포용 폴더 배치
c.batchDist("in_dir", "out_dir", "pw", true, false, null);
//  ↑ forceOutExt: null=입력 확장자 보존, ".hwp" 또는 ".hwpx" 강제 지정 가능

// 개별 파일 N 개 지정 (List<File>)
List<File> files = Arrays.asList(
    new File("a.hwpx"),
    new File("b.hwpx"),
    new File("c.hwpx"));
c.batchFiles(files, "out_dir", ".hwp");

// 개별 파일 배포용 배치
c.batchDistFiles(files, "out_dir", "pw", true, false, null);
```

### 4. 예외 처리

```java
try {
    converter.convertHwpxToHwp(src, dst);
} catch (java.io.IOException e) {
    // 파일 입출력 오류, ZIP 파싱 실패, XXE/zip bomb 차단 등
    System.err.println("변환 실패: " + e.getMessage());
} catch (IllegalArgumentException e) {
    // 입력 == 출력 자기 덮어쓰기 / payload 상한 초과 등 보안 차단
    System.err.println("입력 오류: " + e.getMessage());
} catch (IllegalStateException e) {
    // HWPX 중첩 깊이 상한 초과, 문단 텍스트 상한 초과 등
    System.err.println("악성 입력 의심: " + e.getMessage());
}

// HWP → HWPX 는 hwp2hwpx 내부 예외도 발생 가능 (Exception 포괄)
try {
    converter.convertHwpToHwpx(src, dst);
} catch (IllegalStateException e) {
    // 배포용(DRM) HWP 를 입력한 경우 — 복호화 불가, 원본 HWP 필요
    System.err.println(e.getMessage());
} catch (Exception e) {
    // 기타 변환 오류
    e.printStackTrace();
}
```

### 5. 전체 메서드 시그니처

```java
package kr.n.nframe;

public class HwpConverter {
    // 단건
    public void convertHwpxToHwp(String hwpx, String hwp) throws IOException;
    public void convertHwpToHwpx(String hwp, String hwpx) throws Exception;
    public void makeHwpDist(String in, String out, String password,
                            boolean noCopy, boolean noPrint) throws Exception;

    // 배치 (폴더)
    public BatchResult batchHwpxToHwp(String inDir, String outDir) throws IOException;
    public BatchResult batchHwpToHwpx(String inDir, String outDir) throws IOException;
    public BatchResult batchDist(String inDir, String outDir, String password,
                                 boolean noCopy, boolean noPrint,
                                 String forceOutExt /* null 허용 */) throws IOException;

    // 배치 (파일 리스트)
    public BatchResult batchFiles(List<File> inputs, String outDir,
                                  String toExt /* ".hwp" 또는 ".hwpx" */) throws IOException;
    public BatchResult batchDistFiles(List<File> inputs, String outDir,
                                      String password, boolean noCopy, boolean noPrint,
                                      String forceOutExt) throws IOException;
}
```

`BatchResult` 는 public static 내부 클래스로 다음 필드·getter 를 제공합니다 (v13.11 부터):

```java
HwpConverter.BatchResult r = converter.batchHwpxToHwp("in_dir", "out_dir");
System.out.println("성공: " + r.ok + " / 실패: " + r.fail);
if (r.fail > 0) {
    for (String msg : r.failDetails) {
        System.out.println("  실패: " + msg);
    }
}
```

## 프로젝트 구조

```
src/kr/n/nframe/
├── HwpConverter.java              엔트리포인트 + CLI + 배치
└── hwplib/
    ├── binary/    (RecordWriter, ZlibCompressor, HwpBinaryWriter)
    ├── constants/ (HwpTagId, CtrlId, HwpxNs)
    ├── model/     (HwpDocument 등 순수 값 객체)
    ├── reader/    (HwpxReader, HeaderParser, SectionParser, XmlHelper)
    └── writer/    (HwpWriter, SectionWriter, DocInfoWriter,
                    DistributionWriter, HwpxPostProcessor, HwpxXmlRewriter)
```

## 라이선스
Apache License 2.0 — [LICENSE](LICENSE) 참조.

Copyright © 2026 NSoftware INC.

외부 의존 라이브러리:
- Apache POI (Apache-2.0)
- `kr.dogfoot:hwplib` / `hwpxlib` / `hwp2hwpx` (Apache-2.0 — [github.com/neolord0](https://github.com/neolord0))

## Credits
NSoftware INC. 제공

- **작성** : citadel
- **QA** : 유영진 대리
