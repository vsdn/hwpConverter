# Markdown 샘플 문서

이 문서는 HwpMdConverter 개발/테스트에 사용되는 **대표적인 Markdown 표현**을 모두 담고 있습니다.
GFM(GitHub Flavored Markdown) 기준으로 작성되었습니다.

---

## 1. 제목 (Headings)

# 제목 H1
## 제목 H2
### 제목 H3
#### 제목 H4
##### 제목 H5
###### 제목 H6

---

## 2. 텍스트 강조

- 일반 텍스트입니다.
- **굵은 글씨 (Bold)** 입니다.
- *기울임 (Italic)* 입니다.
- ***굵은 기울임 (Bold + Italic)*** 입니다.
- ~~취소선 (Strikethrough)~~ 입니다.
- `인라인 코드 (Inline code)` 입니다.
- <u>밑줄 (HTML)</u> — Markdown 기본 문법에 없어 HTML 사용.
- <mark>형광펜 (HTML)</mark> — 동일한 이유로 HTML 사용.

---

## 3. 목록

### 3.1 순서 없는 목록

- 항목 1
- 항목 2
  - 하위 항목 2-1
  - 하위 항목 2-2
    - 더 깊은 하위 2-2-1
- 항목 3

### 3.2 순서 있는 목록

1. 첫번째 작업
2. 두번째 작업
   1. 세부 작업 a
   2. 세부 작업 b
3. 세번째 작업

### 3.3 작업 목록 (Task List, GFM)

- [x] 프로젝트 구조 분석
- [x] 표준 문서 수집
- [ ] HwpMdConverter 구현
- [ ] 테스트 작성

---

## 4. 인용문 (Blockquote)

> 이것은 인용문입니다.
> 여러 줄에 걸쳐 작성할 수 있습니다.
>
> > 중첩된 인용문도 가능합니다.
> > — 작성자

---

## 5. 코드 블록

### 5.1 인라인 코드

`printf("Hello, World!");` 또는 `java -jar hwpConverter.jar` 처럼 사용.

### 5.2 코드 펜스 (언어 지정)

```java
public class HelloWorld {
    public static void main(String[] args) {
        System.out.println("Hello, Markdown!");
    }
}
```

```python
def greet(name: str) -> None:
    print(f"Hello, {name}!")

greet("Markdown")
```

```bash
#!/bin/bash
echo "현재 날짜: $(date)"
for file in *.hwp; do
    echo "변환 대상: $file"
done
```

```json
{
    "name": "hwpConverter",
    "version": "1.0.0",
    "features": ["hwp", "hwpx", "markdown"]
}
```

---

## 6. 표 (Table)

| 변환 방향 | 지원 여부 | 비고 |
| --- | :---: | --- |
| HWPX → HWP | ✅ | 내장 구현 |
| HWP → HWPX | ✅ | hwp2hwpx 사용 |
| HWP → DIST | ✅ | AES-128 암호화 |
| HWP → MD | 🚧 | 구현 중 |
| HWPX → MD | 🚧 | 구현 중 |
| MD → HWP | 🚧 | 구현 중 |
| MD → HWPX | 🚧 | 구현 중 |

### 정렬 예시

| 왼쪽 정렬 | 가운데 정렬 | 오른쪽 정렬 |
| :--- | :---: | ---: |
| apple | 100 | $1.00 |
| banana | 50 | $0.50 |
| cherry | 25 | $2.75 |

---

## 7. 링크

- 일반 링크: [Google](https://www.google.com)
- 제목 있는 링크: [Google](https://www.google.com "구글 검색")
- 자동 링크: <https://www.github.com>
- 참조형 링크: [HwpLib GitHub][hwplib]

[hwplib]: https://github.com/neolord0/hwplib "neolord0/hwplib"

---

## 8. 이미지

### 8.1 기본 이미지

![샘플 이미지](sample-image.jpg)

### 8.2 크기/정렬 제어 (HTML)

<p align="center">
  <img src="sample-logo.jpg" alt="샘플 로고" width="100"/>
</p>

### 8.3 링크가 걸린 이미지

[![샘플 이미지 클릭](sample-image.jpg)](https://www.github.com)

---

## 9. 수평선

위 섹션과 아래 섹션을 구분합니다.

---

## 10. 이스케이프 문자

다음 문자는 `\` 로 이스케이프합니다: \* \_ \{ \} \[ \] \( \) \# \+ \- \. \! \| \\ \`

예: 별표를 그대로 표시하려면 `\*별표\*` → \*별표\*

---

## 11. 각주 (Footnote)

HwpConverter는 한글 문서 표준 라이브러리[^hwplib]를 사용합니다.
또한 Apache POI[^poi]도 필요합니다.

[^hwplib]: https://github.com/neolord0/hwplib
[^poi]: Apache POI 5.2.5 - OLE2 Compound Document 처리용

---

## 12. 수식 (Math, 일부 렌더러)

인라인 수식: $E = mc^2$

블록 수식:

$$
\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi}
$$

---

## 13. 펼치기/접기 (HTML)

<details>
<summary>자세한 설치 방법 보기 (클릭)</summary>

1. Java 8+ 설치
2. `hwpConverter.jar` 다운로드
3. 다음 명령 실행:

   ```bash
   java -jar hwpConverter.jar input.hwp output.md
   ```

</details>

---

## 14. 이모지 (일부 렌더러)

:white_check_mark: 성공 &nbsp;
:x: 실패 &nbsp;
:warning: 경고 &nbsp;
:rocket: 배포

---

## 15. 복합 예시 (실전)

다음은 실제 문서에서 자주 쓰이는 혼합 포맷입니다.

### 설치 가이드

> **요구사항**: Java **8** 이상, 디스크 여유 공간 최소 100MB.

1. 다음 파일을 준비하세요:
   - `hwpConverter.jar`
   - `lib/` 폴더 (의존성)
2. 터미널에서 실행:

   ```bash
   java -jar hwpConverter.jar input.hwp output.md
   ```

3. 변환 결과를 확인합니다.

| 오류 코드 | 의미 | 해결 방법 |
| :---: | --- | --- |
| `0` | 정상 | - |
| `2` | 인자 오류 | 사용법 확인 |
| `3` | 변환 실패 | 로그 확인 |

자세한 내용은 [README](../README.md)를 참고하세요.

---

*마지막 업데이트: 2026-04-21*
