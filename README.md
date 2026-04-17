# hwpConverter

HWP · HWPX 양방향 변환기와 한글 배포용(DRM) 문서 생성기.

> **NSoftware INC.**
> 작성: citadel &nbsp;·&nbsp; QA: 유영진 대리

- **HWPX → HWP** : 직접 바이너리 구현 (자체 writer)
- **HWP → HWPX** : `kr.dogfoot:hwplib` / `hwpxlib` / `hwp2hwpx` 기반 + 후처리
- **HWP → 배포용 HWP** : AES-128 암호화 ViewText 스트림 생성 (DRM 플래그 · 복사/인쇄 방지)

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
| HWPX → HWP | `hwpConverter <in.hwpx> <out.hwp>` |
| HWP → HWPX | `hwpConverter <in.hwp> <out.hwpx>` |
| 배포용(DRM) | `hwpConverter --dist <in.hwp> <out.hwp> <password> [--no-copy] [--no-print]` |
| 폴더 배치 | `hwpConverter <inputDir> <outputDir> [--to-hwpx\|--to-hwp]` |
| 파일 지정 | `hwpConverter <f1> <f2> ... <outputDir> --to-hwpx\|--to-hwp` |

경로·암호에 공백이나 셸 특수문자(`< > & |`) 가 있으면 반드시 따옴표로 감싸세요.

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

## 보안 하드닝

입력 검증 · DoS 방어를 기본 적용:

- XML 파싱 시 XXE / DTD / 외부 엔티티 전면 차단
- ZIP 엔트리 개수 · 단일 크기 · 누적 크기 상한 (zip bomb / zip slip 방어)
- `Inflater` 해제 결과 크기 · 압축비 상한 + truncated stream 거부
- 표 / 드로잉 객체 재귀 중첩 깊이 상한
- XML 속성 기반 배열 할당 한도
- 레코드 페이로드 크기 상한 + `Math.addExact` / `Math.multiplyExact`
- 입력 == 출력 경로 자기 덮어쓰기 차단

## 라이선스

Copyright © 2026 NSoftware INC.
Apache License 2.0 — [LICENSE](LICENSE) 참조.

외부 의존 라이브러리:
- Apache POI (Apache-2.0)
- `kr.dogfoot:hwplib` / `hwpxlib` / `hwp2hwpx` (Apache-2.0 — [github.com/neolord0](https://github.com/neolord0))

## Credits

- **작성**: citadel
- **QA**: 유영진 대리
- **소속**: NSoftware INC.
