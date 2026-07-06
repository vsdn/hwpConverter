# Markdown 예제 모음

HwpMdConverter 개발/검증에 쓰이는 **대표적인 Markdown 표현 샘플**과 렌더링용 CSS가 들어 있습니다.

## 파일 목록

| 파일 | 설명 |
| --- | --- |
| `sample.md` | 제목·리스트·표·코드·링크·이미지 등 **모든 자주 쓰는 Markdown 표현**을 포함한 종합 샘플 |
| `preview.html` | `sample.md`를 GitHub 스타일로 렌더링하는 HTML (marked.js 사용) |
| `github-markdown.css` | GitHub 스타일 Markdown 렌더링 CSS (라이트/다크 자동) |
| `github-markdown-light.css` | GitHub 라이트 테마 전용 CSS |
| `github-markdown-dark.css` | GitHub 다크 테마 전용 CSS |
| `sample-image.jpg` | 이미지 포함 테스트용 샘플 그림 (400×200) |
| `sample-logo.jpg` | 소형 로고 샘플 (100×100) |

## 미리보기

CORS 제약으로 로컬 파일(`file://`)에서는 `fetch`가 막히므로 간단한 HTTP 서버를 띄워 확인합니다.

```bash
# 이 폴더에서 실행
python3 -m http.server 8080
# 브라우저에서: http://localhost:8080/preview.html
```

## CSS만 단독 사용하기

다른 마크다운 렌더링 도구에서 GitHub 스타일을 쓰고 싶다면 다음과 같이:

```html
<link rel="stylesheet" href="github-markdown.css">
<article class="markdown-body">
    <!-- 렌더링된 HTML -->
</article>
```

## 출처

- GitHub Markdown CSS: https://github.com/sindresorhus/github-markdown-css (MIT)
- 샘플 이미지: https://picsum.photos (무료)
