package kr.n.nframe.hwplib.writer;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * 방금 작성된 HWPX 파일을 XML 바이트 수준에서 후처리하여 그 구조가
 * 한글 프로그램이 직접 생성하는 것과 더 가깝게 일치하도록 합니다.
 * 목표는 hwpxlib 객체 API로 쉽게 표현할 수 없는 미묘한 레이아웃 측면에 대해
 * 한글이 생성한 HWPX와의 시각적 동등성을 복원하는 것입니다.
 *
 * <p>{@code Contents/section0.xml}에 적용되는 재작성:
 * <ul>
 *   <li><b>말미 빈 run 삽입</b> — 한글은 최소 하나의 텍스트 포함 run이 있는
 *       모든 문단에서 {@code <hp:linesegarray>} 바로 앞에
 *       {@code <hp:run charPrIDRef="0"/>}을 생성합니다. 이는 줄 끝 표시자 역할을 하며
 *       한글이 재배치(reflow) 중에 문단 높이를 계산하는 방식에 영향을 줍니다.
 *       이것이 없으면 거의 빈 페이지의 긴 문단이 축소되고
 *       예상되는 시각적 빈 페이지가 사라집니다
 *       (TA-05의 의도된 빈 2페이지에서 관찰됨).</li>
 *   <li><b>{@code <hp:tc>}의 {@code name=""} 속성</b> — 한글은
 *       모든 표 셀에 빈 {@code name} 속성을 작성합니다.
 *       hwpxlib-1.0.9는 이 속성을 생략합니다. 누락된 속성은 많은 리더에서
 *       허용되지만 엄격한 리더를 혼란스럽게 할 수 있습니다.</li>
 * </ul>
 *
 * <p>또한 한글이 HWPX 출력에 항상 포함하는 {@code META-INF/container.rdf}
 * (매니페스트 RDF)를 추가합니다.
 *
 * <p>구현은 수백 킬로바이트 크기의 섹션 파일에 대해 전체 XML 파싱/직렬화
 * 왕복의 비용과 인코딩 취약성을 피하기 위해 바이트 수준 문자열 치환을
 * 사용합니다.
 */
public final class HwpxXmlRewriter {
    private HwpxXmlRewriter() {}

    private static final String CONTAINER_RDF =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>"
          + "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">"
          + "<rdf:Description rdf:about=\"\">"
          + "<ns0:hasPart xmlns:ns0=\"http://www.hancom.co.kr/hwpml/2016/meta/pkg#\""
          + " rdf:resource=\"Contents/header.xml\"/>"
          + "</rdf:Description>"
          + "<rdf:Description rdf:about=\"Contents/header.xml\">"
          + "<rdf:type rdf:resource=\"http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile\"/>"
          + "</rdf:Description>"
          + "<rdf:Description rdf:about=\"\">"
          + "<ns0:hasPart xmlns:ns0=\"http://www.hancom.co.kr/hwpml/2016/meta/pkg#\""
          + " rdf:resource=\"Contents/section0.xml\"/>"
          + "</rdf:Description>"
          + "<rdf:Description rdf:about=\"Contents/section0.xml\">"
          + "<rdf:type rdf:resource=\"http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile\"/>"
          + "</rdf:Description>"
          + "<rdf:Description rdf:about=\"\">"
          + "<rdf:type rdf:resource=\"http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document\"/>"
          + "</rdf:Description>"
          + "</rdf:RDF>";

    /**
     * {@code path}에 있는 HWPX 파일을 제자리에서 재작성합니다.
     * 어떤 단계라도 실패하면 원본 파일은 그대로 유지됩니다.
     */
    public static void rewrite(String path) {
        rewrite(path, null, null);
    }

    public static void rewrite(String path, List<Integer> astralCodePoints) {
        rewrite(path, astralCodePoints, null);
    }

    /**
     * 선택적 astral 코드 포인트 목록과 탭 설정 목록으로 재작성합니다.
     * <ul>
     *  <li>{@code astralCodePoints} — hwplib의 {@code HWPCharNormal.getCh()}가
     *      잃어버린 surrogate pair 문자를 복구합니다.</li>
     *  <li>{@code tabSettings} — 소스 순서대로 탭별
     *      {@code int[]{width,leader,type}} 목록. hwp2hwpx-1.0.0의
     *      {@code ForChars.addTab()}이 {@code width=4000 leader=0(NONE) type=1(LEFT)}로
     *      하드코딩한 원래 탭 매개변수 (목차 점 리더, 우측 정렬 페이지 번호 탭 등)를
     *      복원합니다. 소스 HWP에서 목록을 채우는 방법은
     *      {@link kr.n.nframe.HwpConverter#convertHwpToHwpx}를 참조하십시오.</li>
     * </ul>
     */
    public static void rewrite(String path, List<Integer> astralCodePoints,
                               List<int[]> tabSettings) {
        try {
            rewriteUnchecked(path, astralCodePoints, tabSettings);
        } catch (Throwable t) {
            System.err.println("[HwpxXmlRewriter] Skipping (" + t.getClass().getSimpleName()
                    + ": " + t.getMessage() + ")");
        }
    }

    /** ZIP 엔트리 최대 개수. */
    private static final int MAX_ENTRIES = 10_000;
    /** 단일 엔트리 최대 크기. */
    private static final long MAX_ENTRY_BYTES = 64L * 1024 * 1024;
    /** 모든 엔트리 누적 최대 크기. */
    private static final long MAX_TOTAL_BYTES = 256L * 1024 * 1024;

    private static void rewriteUnchecked(String path, List<Integer> astralCodePoints,
                                         List<int[]> tabSettings) throws IOException {
        Path src = Paths.get(path);
        // 모든 엔트리를 메모리로 읽습니다. 엔트리 개수·단일/합계 크기 상한을 적용해
        // 악성/손상 ZIP 에 의한 OOM(zip bomb) 을 방지합니다.
        LinkedHashMap<String, byte[]> entries = new LinkedHashMap<>();
        long totalBytes = 0L;
        int entryCount = 0;
        try (ZipInputStream zin = new ZipInputStream(Files.newInputStream(src))) {
            ZipEntry e;
            byte[] buf = new byte[8192];
            while ((e = zin.getNextEntry()) != null) {
                if (++entryCount > MAX_ENTRIES) {
                    throw new IOException("HWPX entry count exceeds limit (" + MAX_ENTRIES + ")");
                }
                validateEntryName(e.getName());
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                int n;
                long entryBytes = 0;
                while ((n = zin.read(buf)) > 0) {
                    entryBytes += n;
                    if (entryBytes > MAX_ENTRY_BYTES) {
                        throw new IOException("HWPX entry too large: " + e.getName());
                    }
                    if (totalBytes + entryBytes > MAX_TOTAL_BYTES) {
                        throw new IOException("HWPX total size exceeds limit ("
                                + MAX_TOTAL_BYTES + " bytes)");
                    }
                    out.write(buf, 0, n);
                }
                totalBytes += entryBytes;
                entries.put(e.getName(), out.toByteArray());
                zin.closeEntry();
            }
        }

        // section0.xml 재작성
        byte[] sec = entries.get("Contents/section0.xml");
        if (sec != null) {
            String xml = new String(sec, StandardCharsets.UTF_8);
            // 1) surrogate pair (astral) 문자 복구
            String step1 = recoverAstralChars(xml, astralCodePoints);
            // 2) 원래 탭 매개변수 복원
            String step2 = recoverTabSettings(step1, tabSettings);
            // 3) 구조적 재작성 적용 (tc name, 말미 빈 run)
            String rewritten = rewriteSection(step2);
            if (!rewritten.equals(xml)) {
                entries.put("Contents/section0.xml", rewritten.getBytes(StandardCharsets.UTF_8));
            }
        }

        // 누락된 경우 container.rdf 추가
        if (!entries.containsKey("META-INF/container.rdf")) {
            entries.put("META-INF/container.rdf", CONTAINER_RDF.getBytes(StandardCharsets.UTF_8));
        }

        // 임시 파일로 다시 작성한 후 원본을 원자적으로 대체합니다.
        // 예외 경로에서 tmp 가 남지 않도록 try/finally 로 정리.
        Path tmp = src.resolveSibling(src.getFileName() + ".rewrite.tmp");
        boolean moved = false;
        try {
            try (ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(tmp.toFile()))) {
                for (Map.Entry<String, byte[]> me : entries.entrySet()) {
                    ZipEntry ze = new ZipEntry(me.getKey());
                    zout.putNextEntry(ze);
                    zout.write(me.getValue());
                    zout.closeEntry();
                }
            }
            Files.move(tmp, src, StandardCopyOption.REPLACE_EXISTING);
            moved = true;
        } finally {
            if (!moved) {
                try { Files.deleteIfExists(tmp); } catch (IOException ignored) {}
            }
        }
    }

    /**
     * ZIP 엔트리 이름에 traversal / 절대경로 / NUL byte / 드라이브 레터가 있는지 검사.
     * 메모리 상 map 에만 저장하는 현 구조에서도 향후 디스크 추출 확장에 대한
     * 이중 방어로 엄격히 거부한다.
     */
    private static void validateEntryName(String name) throws IOException {
        if (name == null) throw new IOException("Null ZIP entry name");
        if (name.indexOf('\0') >= 0) throw new IOException("Malicious ZIP entry name (NUL byte)");
        String norm = name.replace('\\', '/');
        if (norm.startsWith("/")
                || norm.contains("../")
                || norm.equals("..")
                || norm.endsWith("/..")
                || (norm.length() >= 2 && norm.charAt(1) == ':')) {
            throw new IOException("Malicious ZIP entry path: " + name);
        }
    }

    /**
     * 말미 빈 run과 표 셀에 {@code name=""}을 추가합니다.
     * 순수 문자열 처리 — 파일의 나머지 부분을 바이트 단위로 정확하게 유지합니다.
     */
    /**
     * hwplib의 {@code HWPCharNormal.getCh()}가 U+FFFD U+FFFD로 훼손한
     * surrogate pair (astral-plane) 문자를 복구합니다.
     *
     * <p>근본 원인: hwplib는 각 HWP UINT16 코드 유닛을 별도의 2-byte UTF-16LE
     * String으로 디코딩합니다. HWP 소스에 surrogate pair (예: U+F02B1과 같은
     * Private Use Area 코드 포인트)가 포함되어 있을 때, 쌍의 각 절반은
     * 잘못 형성된 독립 String이 되고 Java는 이를 U+FFFD (유니코드 대체 문자)로
     * 교체합니다.
     *
     * <p>전략: 호출자가 소스 HWP를 순차적으로 스캔하여 발견한 모든 astral 코드
     * 포인트 (U+10000..U+10FFFF)를 수집했습니다. 여기서는 {@code <hp:t>} 내용의
     * 모든 {@code U+FFFDU+FFFD} 출현을 해당 목록의 다음 astral 코드 포인트로
     * 순차적으로 교체하여 올바른 surrogate pair를 생성합니다.
     *
     * @param xml                 section0.xml 내용
     * @param astralCodePoints    소스 순서의 코드 포인트, 건너뛰려면 null
     */
    static String recoverAstralChars(String xml, List<Integer> astralCodePoints) {
        if (astralCodePoints == null || astralCodePoints.isEmpty()) return xml;
        // 연속된 두 개의 U+FFFD 문자가 훼손된 surrogate pair 서명을 형성합니다.
        StringBuilder out = new StringBuilder(xml.length());
        int i = 0, cpIdx = 0, n = xml.length();
        while (i < n) {
            char c = xml.charAt(i);
            if (c == 0xFFFD && i + 1 < n && xml.charAt(i + 1) == 0xFFFD
                    && cpIdx < astralCodePoints.size()) {
                int cp = astralCodePoints.get(cpIdx++);
                out.appendCodePoint(cp);
                i += 2;
            } else {
                out.append(c);
                i++;
            }
        }
        return out.toString();
    }

    /**
     * HWP leader byte → HWPX {@code <hp:tab leader="...">} enum 이름.
     * 값은 HWP 5.0 스펙 §tab 데이터 및 hwpxlib LineType2 enum에서 가져옴.
     */
    private static String leaderName(int leader) {
        switch (leader) {
            case 0: return "NONE";
            case 1: return "SOLID";
            case 2: return "DASH";
            case 3: return "DOT";
            case 4: return "DASH_DOT";
            case 5: return "DASH_DOT_DOT";
            case 6: return "LONG_DASH";
            case 7: return "CIRCLE";
            case 8: return "DOUBLE_SLIM";
            case 9: return "SLIM_THICK";
            case 10: return "THICK_SLIM";
            case 11: return "SLIM_THICK_SLIM";
            default: return "NONE";
        }
    }

    /**
     * HWP tab type → HWPX {@code <hp:tab type="...">} enum 이름.
     * 0=LEFT, 1=RIGHT, 2=CENTER, 3=DECIMAL.
     */
    private static String tabTypeName(int type) {
        switch (type) {
            case 0: return "LEFT";
            case 1: return "RIGHT";
            case 2: return "CENTER";
            case 3: return "DECIMAL";
            default: return "LEFT";
        }
    }

    /**
     * hwp2hwpx-1.0.0의 {@code ForChars.addTab()}이 덮어쓴 탭별
     * width/leader/type 매개변수를 복원합니다. XML의 모든
     * {@code <hp:tab .../>}를 {@code tabSettings}의 다음 요소로부터 구성된
     * 것으로 순차적으로 교체합니다.
     *
     * <p>HWPX 숫자 enum 매핑 (hwpxlib 직렬화):
     * <ul>
     *   <li>leader: LineType2 — 0 NONE, 1 SOLID, 2 DASH, 3 DOT, 4 DASH_DOT,
     *       5 DASH_DOT_DOT, 6 LONG_DASH, 등.</li>
     *   <li>type: TabItemType — 0 LEFT, 1 RIGHT, 2 CENTER, 3 DECIMAL.</li>
     * </ul>
     * hwpxlib writer는 enum의 {@code str()}이 숫자를 반환할 때 숫자 속성을
     * 생성합니다 (한글이 생성하는 {@code leader="3" type="2"} 형식) —
     * LineType2 및 TabItemType에 대해 그렇게 합니다. 따라서 숫자를 직접
     * 생성합니다.
     */
    static String recoverTabSettings(String xml, List<int[]> tabSettings) {
        if (tabSettings == null || tabSettings.isEmpty()) return xml;
        StringBuilder out = new StringBuilder(xml.length());
        final String TAB_PREFIX = "<hp:tab ";
        int i = 0, idx = 0, n = xml.length();
        while (i < n) {
            int p = xml.indexOf(TAB_PREFIX, i);
            if (p < 0) { out.append(xml, i, n); break; }
            out.append(xml, i, p);
            int end = xml.indexOf("/>", p);
            if (end < 0) { out.append(xml, p, n); break; }
            if (idx < tabSettings.size()) {
                int[] t = tabSettings.get(idx++);
                out.append("<hp:tab width=\"").append(t[0])
                   .append("\" leader=\"").append(t[1])
                   .append("\" type=\"").append(t[2])
                   .append("\"/>");
            } else {
                // 설정이 소진됨 — 기존 태그를 변경 없이 복사
                out.append(xml, p, end + 2);
            }
            i = end + 2;
        }
        return out.toString();
    }

    static String rewriteSection(String xml) {
        // 단일 패스 스캔 — 큰 HWPX 파일 (100MB+)에 필수. XML을 처음부터
        // 끝까지 순회하면서 subList 중첩 깊이, 문단-텍스트-존재 플래그,
        // 마지막으로 본 `</hp:run>` 위치를 추적하고 한 번의 패스로
        // 재작성을 생성합니다.
        //
        // 적용되는 재작성:
        //   · subList 내의 모든 `<hp:linesegarray>` 앞에
        //     `<hp:run charPrIDRef="0"/>` 삽입 (이전 요소가 `</hp:run>`이고,
        //     문단에 텍스트가 포함되어 있으며, 빈 run이 앞에 오지
        //     않는 경우).
        //   · `name=` 속성이 없는 모든 `<hp:tc ...>` 여는 태그에
        //     `name=""` 주입.
        final String SUB_OPEN_S  = "<hp:subList";
        final String SUB_CLOSE_S = "</hp:subList>";
        final String LSA        = "<hp:linesegarray>";
        final String END_RUN    = "</hp:run>";
        final String EMPTY_RUN  = "<hp:run charPrIDRef=\"0\"/>";
        final String P_OPEN     = "<hp:p ";
        final String T_OPEN     = "<hp:t>";
        final String TC_OPEN    = "<hp:tc ";

        StringBuilder out = new StringBuilder(xml.length() + 64 * 1024);
        int subDepth = 0;
        int lastPOpenOut = -1;   // 현재 문단이 시작된 OUTPUT에서의 offset
        int lastEndRunOut = -1;  // 마지막 </hp:run>이 끝난 OUTPUT에서의 offset
        boolean paraHasText = false;

        int i = 0;
        int n = xml.length();
        while (i < n) {
            char c = xml.charAt(i);
            if (c != '<') {
                out.append(c);
                i++;
                continue;
            }
            // 태그 시작.
            if (xml.regionMatches(i, SUB_CLOSE_S, 0, SUB_CLOSE_S.length())) {
                out.append(SUB_CLOSE_S);
                i += SUB_CLOSE_S.length();
                if (subDepth > 0) subDepth--;
            } else if (xml.regionMatches(i, SUB_OPEN_S, 0, SUB_OPEN_S.length())) {
                // <hp:subList ...>는 /> 또는 >로 닫힐 수 있음. 그대로 복사.
                int end = xml.indexOf('>', i);
                if (end < 0) { out.append(xml, i, n); break; }
                out.append(xml, i, end + 1);
                // self-closing <.../>은 깊이를 증가시키지 않음; 그렇지 않으면 증가시킴.
                boolean selfClosing = xml.charAt(end - 1) == '/';
                if (!selfClosing) subDepth++;
                i = end + 1;
            } else if (xml.regionMatches(i, P_OPEN, 0, P_OPEN.length())) {
                int end = xml.indexOf('>', i);
                if (end < 0) { out.append(xml, i, n); break; }
                out.append(xml, i, end + 1);
                lastPOpenOut = out.length();
                lastEndRunOut = -1;
                paraHasText = false;
                i = end + 1;
            } else if (xml.regionMatches(i, END_RUN, 0, END_RUN.length())) {
                out.append(END_RUN);
                lastEndRunOut = out.length();
                i += END_RUN.length();
            } else if (xml.regionMatches(i, T_OPEN, 0, T_OPEN.length())) {
                out.append(T_OPEN);
                paraHasText = true;
                i += T_OPEN.length();
            } else if (xml.regionMatches(i, LSA, 0, LSA.length())) {
                // 조건이 충족되면 말미 빈 run 삽입
                if (subDepth > 0 && paraHasText && lastEndRunOut == out.length()) {
                    // 중복 방지: 출력에서 앞에 있는 내용 확인
                    int lo = out.length() - EMPTY_RUN.length();
                    boolean dup = lo >= 0 && out.indexOf(EMPTY_RUN, lo) == lo;
                    if (!dup) out.append(EMPTY_RUN);
                }
                out.append(LSA);
                i += LSA.length();
            } else if (xml.regionMatches(i, TC_OPEN, 0, TC_OPEN.length())) {
                int end = xml.indexOf('>', i);
                if (end < 0) { out.append(xml, i, n); break; }
                String tag = xml.substring(i, end + 1);
                if (!tag.contains(" name=\"")) {
                    out.append(TC_OPEN).append("name=\"\" ").append(tag.substring(TC_OPEN.length()));
                } else {
                    out.append(tag);
                }
                i = end + 1;
            } else {
                out.append(c);
                i++;
            }
        }
        return out.toString();
    }

    /**
     * 휴리스틱: offset {@code pos}를 감싸는 문단에 비어 있지 않은
     * {@code <hp:t>} 요소가 하나 이상 포함되어 있습니까? 체크 범위를
     * 정하기 위해 가장 가까운 {@code <hp:p }까지 거꾸로 검색합니다.
     */
    private static boolean paragraphHasText(String xml, int pos) {
        int pStart = xml.lastIndexOf("<hp:p ", pos);
        if (pStart < 0) return false;
        int scan = xml.indexOf("<hp:t>", pStart);
        // 이 문단을 감싸는 <hp:run/> 체인 내에 있어야 합니다.
        return scan >= 0 && scan < pos;
    }

    /**
     * {@code pos}의 문단이 {@code <hp:subList>} 안에 있습니까?
     * subList는 표 셀 내용과 머리말/꼬리말 본문을 감싸며 — 한글이
     * 말미 빈 run을 추가하는 컨텍스트입니다.
     * 문단 시작부터 거꾸로 감싸는 subList 여는/닫는 태그를 세어
     * 감지합니다.
     */
    private static boolean isInsideSubList(String xml, int pos) {
        // pos 이전의 <hp:subList> 및 </hp:subList> 태그를 세기; open > close이면 내부에 있음.
        int depth = 0;
        int i = 0;
        while (i < pos) {
            int openPos = xml.indexOf("<hp:subList", i);
            int closePos = xml.indexOf("</hp:subList>", i);
            if (openPos < 0 && closePos < 0) break;
            if (openPos < 0 || (closePos >= 0 && closePos < openPos)) {
                if (closePos >= pos) break;
                depth--;
                i = closePos + 13;
            } else {
                if (openPos >= pos) break;
                depth++;
                i = openPos + 11;
            }
        }
        return depth > 0;
    }
}
