import kr.n.nframe.HwpConverter;
import kr.n.nframe.HwpConverter.BatchResult;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

/**
 * hwpConverter.jar 의 모든 public API 를 호출해 보는 통합 테스트.
 *
 *  - 단건  3종 : HWPX→HWP, HWP→HWPX, 배포용(DRM)
 *  - 배치  4종 : 폴더 HWPX→HWP, 폴더 HWP→HWPX, 폴더 배포용, 파일목록 변환
 *
 * 실행:
 *   javac -cp "../hwpConverter.jar;../lib/*" -d build HwpConverterApiTest.java
 *   java  -cp "build;../hwpConverter.jar;../lib/*" HwpConverterApiTest
 *
 * 또는 apitest/run.bat 사용.
 */
public class HwpConverterApiTest {

    private static final String INPUT_DIR = "../hwp/pages";
    private static final String OUT_DIR   = "out";

    private static int passed = 0;
    private static int failed = 0;
    private static final List<String> failures = new ArrayList<>();

    public static void main(String[] args) throws Exception {
        bannerHeader();

        // 입력 파일 존재 확인
        if (!new File(INPUT_DIR, "00.hwpx").exists() || !new File(INPUT_DIR, "00.hwp").exists()) {
            System.err.println("[ERROR] 테스트 입력 파일이 없습니다: " + INPUT_DIR);
            System.exit(2);
        }

        // 출력 디렉토리 정리
        cleanDir(OUT_DIR);
        Files.createDirectories(Paths.get(OUT_DIR));

        HwpConverter conv = new HwpConverter();

        // ─── 단건 변환 ────────────────────────────────────────────────
        section("Single-file API");

        testCase("convertHwpxToHwp(00.hwpx → out/single_00.hwp)", () -> {
            conv.convertHwpxToHwp(INPUT_DIR + "/00.hwpx", OUT_DIR + "/single_00.hwp");
            assertNonEmpty(OUT_DIR + "/single_00.hwp");
        });

        testCase("convertHwpToHwpx(00.hwp → out/single_00.hwpx)", () -> {
            conv.convertHwpToHwpx(INPUT_DIR + "/00.hwp", OUT_DIR + "/single_00.hwpx");
            assertNonEmpty(OUT_DIR + "/single_00.hwpx");
        });

        testCase("makeHwpDist(00.hwp + pw → out/single_00_dist.hwp)", () -> {
            conv.makeHwpDist(INPUT_DIR + "/00.hwp",
                             OUT_DIR + "/single_00_dist.hwp",
                             "test1234",
                             false,   // noCopy
                             false);  // noPrint
            assertNonEmpty(OUT_DIR + "/single_00_dist.hwp");
        });

        // ─── 배치 변환 ────────────────────────────────────────────────
        section("Batch API");

        testBatch("batchHwpxToHwp(폴더 전체 HWPX→HWP)", () ->
            conv.batchHwpxToHwp(INPUT_DIR, OUT_DIR + "/batch_to_hwp"));

        testBatch("batchHwpToHwpx(폴더 전체 HWP→HWPX)", () ->
            conv.batchHwpToHwpx(INPUT_DIR, OUT_DIR + "/batch_to_hwpx"));

        testBatch("batchDist(폴더 전체 + 비밀번호)", () ->
            conv.batchDist(INPUT_DIR, OUT_DIR + "/batch_dist",
                           "pw1234", false, false));

        // 파일 목록 (앞 3개의 .hwpx)
        File[] hwpxList = new File(INPUT_DIR).listFiles((d, n) -> n.endsWith(".hwpx"));
        if (hwpxList != null && hwpxList.length > 0) {
            Arrays.sort(hwpxList);
            List<File> picked = new ArrayList<>(
                    Arrays.asList(Arrays.copyOf(hwpxList, Math.min(3, hwpxList.length))));
            testBatch("batchFiles(파일 목록 HWPX→HWP, " + picked.size() + "건)", () ->
                conv.batchFiles(picked, OUT_DIR + "/batch_files", "hwp"));
        }

        // ─── 결과 ─────────────────────────────────────────────────────
        bannerFooter();
    }

    // ─── 테스트 유틸 ────────────────────────────────────────────────

    @FunctionalInterface
    interface VoidCall { void run() throws Exception; }

    @FunctionalInterface
    interface BatchCall { BatchResult run() throws Exception; }

    private static void testCase(String name, VoidCall call) {
        long t0 = System.currentTimeMillis();
        try {
            call.run();
            System.out.printf("  [PASS] %-58s (%4d ms)%n",
                              name, System.currentTimeMillis() - t0);
            passed++;
        } catch (Throwable e) {
            System.out.printf("  [FAIL] %-58s%n           → %s%n",
                              name, e.toString());
            failures.add(name + " : " + e);
            failed++;
        }
    }

    private static void testBatch(String name, BatchCall call) {
        long t0 = System.currentTimeMillis();
        try {
            BatchResult r = call.run();
            System.out.printf("  [PASS] %-58s ok=%d fail=%d (%4d ms)%n",
                              name, r.ok, r.fail, System.currentTimeMillis() - t0);
            if (r.fail > 0) {
                for (String d : r.failDetails) {
                    System.out.println("           · " + d);
                }
            }
            passed++;
        } catch (Throwable e) {
            System.out.printf("  [FAIL] %-58s%n           → %s%n",
                              name, e.toString());
            failures.add(name + " : " + e);
            failed++;
        }
    }

    private static void assertNonEmpty(String path) {
        File f = new File(path);
        if (!f.exists())            throw new AssertionError("출력 파일 없음: " + path);
        if (f.length() == 0)        throw new AssertionError("출력 파일 0 byte: " + path);
    }

    private static void cleanDir(String dir) throws Exception {
        Path p = Paths.get(dir);
        if (!Files.exists(p)) return;
        Files.walk(p)
             .sorted(Comparator.reverseOrder())
             .forEach(f -> { try { Files.delete(f); } catch (Exception ignore) {} });
    }

    private static void bannerHeader() {
        System.out.println();
        System.out.println("┌──────────────────────────────────────────────────────────────────────┐");
        System.out.println("│  hwpConverter API 통합 테스트                                        │");
        System.out.println("└──────────────────────────────────────────────────────────────────────┘");
    }

    private static void section(String label) {
        StringBuilder sb = new StringBuilder("── ").append(label).append(' ');
        int n = Math.max(2, 64 - label.length());
        for (int i = 0; i < n; i++) sb.append('─');
        System.out.println();
        System.out.println(sb.toString());
    }

    private static void bannerFooter() {
        System.out.println();
        System.out.println("──────────────────────────────────────────────────────────────────────");
        System.out.printf ("  결과: %d passed, %d failed%n", passed, failed);
        System.out.println("──────────────────────────────────────────────────────────────────────");
        if (failed > 0) {
            System.out.println("\n실패 상세:");
            for (String f : failures) System.out.println("  - " + f);
            System.exit(1);
        }
    }
}
