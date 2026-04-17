package kr.n.nframe.hwplib.writer;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.docinfo.DocInfo;
import kr.dogfoot.hwplib.object.docinfo.compatibledocument.CompatibleDocumentSort;
import kr.dogfoot.hwpxlib.object.HWPXFile;
import kr.dogfoot.hwpxlib.object.common.ObjectList;
import kr.dogfoot.hwpxlib.object.content.header_xml.CompatibleDocument;
import kr.dogfoot.hwpxlib.object.content.header_xml.HeaderXMLFile;
import kr.dogfoot.hwpxlib.object.content.header_xml.LayoutCompatibilityItem;
import kr.dogfoot.hwpxlib.object.content.header_xml.RefList;
import kr.dogfoot.hwpxlib.object.content.header_xml.enumtype.HorizontalAlign1;
import kr.dogfoot.hwpxlib.object.content.header_xml.enumtype.TargetProgramSort;
import kr.dogfoot.hwpxlib.object.content.header_xml.references.Numbering;
import kr.dogfoot.hwpxlib.object.content.header_xml.references.numbering.ParaHead;
import kr.dogfoot.hwpxlib.object.content.section_xml.ParaListCore;
import kr.dogfoot.hwpxlib.object.content.section_xml.SectionXMLFile;
import kr.dogfoot.hwpxlib.object.content.section_xml.SubList;
import kr.dogfoot.hwpxlib.object.content.section_xml.enumtype.TablePageBreak;
import kr.dogfoot.hwpxlib.object.content.section_xml.paragraph.Para;
import kr.dogfoot.hwpxlib.object.content.section_xml.paragraph.Run;
import kr.dogfoot.hwpxlib.object.content.section_xml.paragraph.RunItem;
import kr.dogfoot.hwpxlib.object.content.section_xml.paragraph.object.Table;
import kr.dogfoot.hwpxlib.object.content.section_xml.paragraph.object.table.Tc;
import kr.dogfoot.hwpxlib.object.content.section_xml.paragraph.object.table.Tr;

/**
 * {@code Hwp2Hwpx.toHWPX}가 생성한 {@link HWPXFile}의 후처리기.
 *
 * <p>목표: 한글 워드프로세서가 동일 HWP를 HWPX로 저장했을 때의 결과와
 * 변환기 출력을 일치시켜서, 한글에서 열었을 때 동일하게 렌더링/동작하도록 한다.
 *
 * <p><b>한글의 동작은 단일 변환이 아니다 — 원본 HWP의
 * {@code HWPTAG_COMPATIBLE_DOCUMENT.targetProgram}에 따라 달라진다.</b>
 * 실측 결과 (한글 HWP 2024 기준):
 *
 * <pre>
 *   HWP 원본 targetProgram    →  HWPX 출력
 *   ────────────────────────     ──────────────────────────────────────
 *   HWPCurrent  (0)           →  targetProgram="MS_WORD"
 *                                + layoutCompatibility 플래그 35개
 *                                + 표 pageBreak → CELL
 *                                + paraHead align LEFT
 *
 *   MSWord      (2)           →  targetProgram="HWP201X"
 *                                + 빈 &lt;hh:layoutCompatibility/&gt;
 *                                + 표는 pageBreak=TABLE 유지 (hwp2hwpx 기본값)
 *                                + paraHead align 유지
 *
 *   HWP2007     (1)           →  (참조 파일 없음; hwp2hwpx 기본값 유지)
 * </pre>
 *
 * <p>이러한 비대칭 매핑은 정확한 페이지 레이아웃 렌더링을 위해 중요하다:
 * MS-Word 호환 모드로 작성된 문서에 MS_WORD + 35 플래그를 강제하면
 * 문단 높이 재계산으로 인해 빈 페이지가 사라질 수 있고 (예: TA-05의
 * 의도된 빈 페이지 2) 요소 간격이 변할 수 있다. 네이티브 문서에
 * HWP201X 빈 값을 강제하면 표와 문단이 밀착되어 렌더링된다.
 * 한글은 원본 값에 따라 분기하여 두 경우를 모두 회피한다.
 *
 * @see #normalize(HWPFile, HWPXFile)
 */
public final class HwpxPostProcessor {
    /**
     * {@code targetProgram == HWPCurrent}인 HWP 원본에 대해 한글이
     * {@code <hh:layoutCompatibility>} 아래에 출력하는 태그 이름들.
     * hwpxlib의 writer가 이 이름들을 그대로 기록하므로, 한글과 byte 단위로
     * 일치시키기 위해 {@code hh:} 네임스페이스 접두사도 여기에 포함한다.
     */
    private static final String[] LAYOUT_COMPAT_FLAGS = {
            "hh:applyFontWeightToBold",
            "hh:useInnerUnderline",
            "hh:useLowercaseStrikeout",
            "hh:extendLineheightToOffset",
            "hh:treatQuotationAsLatin",
            "hh:doNotAlignWhitespaceOnRight",
            "hh:doNotAdjustWordInJustify",
            "hh:baseCharUnitOnEAsian",
            "hh:baseCharUnitOfIndentOnFirstChar",
            "hh:adjustLineheightToFont",
            "hh:adjustBaselineInFixedLinespacing",
            "hh:applyPrevspacingBeneathObject",
            "hh:applyNextspacingOfLastPara",
            "hh:adjustParaBorderfillToSpacing",
            "hh:connectParaBorderfillOfEqualBorder",
            "hh:adjustParaBorderOffsetWithBorder",
            "hh:extendLineheightToParaBorderOffset",
            "hh:applyParaBorderToOutside",
            "hh:applyMinColumnWidthTo1mm",
            "hh:applyTabPosBasedOnSegment",
            "hh:breakTabOverLine",
            "hh:adjustVertPosOfLine",
            "hh:doNotAlignLastForbidden",
            "hh:adjustMarginFromAdjustLineheight",
            "hh:baseLineSpacingOnLineGrid",
            "hh:applyCharSpacingToCharGrid",
            "hh:doNotApplyGridInHeaderFooter",
            "hh:applyExtendHeaderFooterEachSection",
            "hh:doNotApplyLinegridAtNoLinespacing",
            "hh:doNotAdjustEmptyAnchorLine",
            "hh:overlapBothAllowOverlap",
            "hh:extendVertLimitToPageMargins",
            "hh:doNotHoldAnchorOfTable",
            "hh:doNotFormattingAtBeneathAnchor",
            "hh:adjustBaselineOfObjectToBottom",
    };

    private HwpxPostProcessor() {}

    /**
     * 원본 HWP의 호환성 설정에 따라 한글과 일치하는 정규화를 적용한다.
     */
    public static void normalize(HWPFile hwp, HWPXFile hwpx) {
        if (hwpx == null) return;
        CompatibleDocumentSort srcSort = sourceTargetProgram(hwp);
        if (srcSort == CompatibleDocumentSort.HWPCurrent) {
            applyHwpCurrentProfile(hwpx);
        } else if (srcSort == CompatibleDocumentSort.MSWord) {
            applyMsWordProfile(hwpx);
        }
        // HWP2007: 참조 데이터가 없어 의도적으로 건드리지 않는다.
    }

    private static CompatibleDocumentSort sourceTargetProgram(HWPFile hwp) {
        if (hwp == null) return null;
        DocInfo di = hwp.getDocInfo();
        if (di == null || di.getCompatibleDocument() == null) return null;
        return di.getCompatibleDocument().getTargetProgream();
    }

    // ------------------------------------------------------------------
    // 프로파일: HWP 원본 targetProgram == HWPCurrent
    // (한글은 MS_WORD + 35 플래그, CELL 표, LEFT paraHead로 출력)
    // ------------------------------------------------------------------

    private static void applyHwpCurrentProfile(HWPXFile hwpx) {
        // HWPCurrent 원본 문서에 대한 프로파일 선택은 1:1이 아니다.
        // 실측 규칙 (한글 12.0.0.535):
        //   · hwp2hwpx가 header.xml에 <hh:memoPr>을 하나 이상 출력하면
        //     (즉, HWP 원본에 메모 모양이 선언된 경우) 한글은 MS_WORD와
        //     35 플래그 전체 layoutCompatibility 블록으로 저장한다.
        //     (TA-05_100p가 이 경우에 해당 — 제목의 "색 채우기 범위"가
        //     이 플래그들에 의존한다.)
        //   · 그 외에는 hwp2hwpx의 기본값(HWP201X + 빈 값)을 유지한다.
        //     이렇게 해야 의도된 빈 페이지(TA-05_2p), 페이지 수가 올바르게
        //     보존되며 메모 없는 문서는 일반적으로 한글 출력과 일치한다.
        if (hasMemoShapes(hwpx)) {
            setCompatibleDocument(hwpx, TargetProgramSort.MS_WORD, LAYOUT_COMPAT_FLAGS);
        }
        fixParaHeadAlignment(hwpx);
        remapTablePageBreak(hwpx, TablePageBreak.TABLE, TablePageBreak.CELL);
    }

    private static boolean hasMemoShapes(HWPXFile hwpx) {
        HeaderXMLFile hdr = hwpx.headerXMLFile();
        if (hdr == null) return false;
        RefList rl = hdr.refList();
        if (rl == null) return false;
        try {
            ObjectList<kr.dogfoot.hwpxlib.object.content.header_xml.references.MemoPr> memo = rl.memoProperties();
            return memo != null && memo.count() > 0;
        } catch (Throwable t) {
            return false;
        }
    }

    // ------------------------------------------------------------------
    // 프로파일: HWP 원본 targetProgram == MSWord
    // (한글은 HWP201X + 빈 layoutCompatibility로 출력, 기본값 유지)
    // ------------------------------------------------------------------

    private static void applyMsWordProfile(HWPXFile hwpx) {
        // targetProgram=MSWord인 HWP 원본은 한글에서 두 가지로 갈린다:
        //   · 메모 모양 있음    → MS_WORD + 35 플래그 + CELL page-break
        //     (예: TA-05_100p — 제목의 "색 채우기 범위"/표 pageBreak가
        //      이들에 의존한다.)
        //   · 메모 모양 없음    → HWP201X + 빈 값, 기본 pageBreak 유지
        //     (예: TA-05_100p_2p — 의도된 빈 페이지 2가 그렇지 않으면
        //      사라진다.)
        if (hasMemoShapes(hwpx)) {
            setCompatibleDocument(hwpx, TargetProgramSort.MS_WORD, LAYOUT_COMPAT_FLAGS);
            remapTablePageBreak(hwpx, TablePageBreak.TABLE, TablePageBreak.CELL);
        } else {
            setCompatibleDocument(hwpx, TargetProgramSort.HWP201X, new String[0]);
        }
    }

    // ------------------------------------------------------------------
    // 공용 헬퍼
    // ------------------------------------------------------------------

    private static void setCompatibleDocument(HWPXFile hwpx,
                                              TargetProgramSort target,
                                              String[] flags) {
        HeaderXMLFile header = hwpx.headerXMLFile();
        if (header == null) return;
        if (header.compatibleDocument() == null) header.createCompatibleDocument();
        CompatibleDocument cd = header.compatibleDocument();
        cd.targetProgram(target);

        if (cd.layoutCompatibility() == null) cd.createLayoutCompatibility();
        ObjectList<LayoutCompatibilityItem> list = cd.layoutCompatibility();

        // 기존 항목을 모두 제거 (hwp2hwpx가 추가했을 수 있음) 후 다시 추가.
        list.removeAll();
        for (String flag : flags) {
            LayoutCompatibilityItem item = list.addNew();
            item.name(flag);
        }
    }

    private static void fixParaHeadAlignment(HWPXFile hwpx) {
        HeaderXMLFile header = hwpx.headerXMLFile();
        if (header == null) return;
        RefList refs = header.refList();
        if (refs == null) return;
        ObjectList<Numbering> numberings = refs.numberings();
        if (numberings == null) return;
        for (int i = 0; i < numberings.count(); i++) {
            Numbering n = numberings.get(i);
            int phCount = n.countOfParaHead();
            for (int k = 0; k < phCount; k++) {
                ParaHead ph = n.getParaHead(k);
                if (ph.align() == HorizontalAlign1.CENTER) {
                    ph.align(HorizontalAlign1.LEFT);
                }
            }
        }
    }

    private static void remapTablePageBreak(HWPXFile hwpx,
                                            TablePageBreak from,
                                            TablePageBreak to) {
        int sections = hwpx.sectionXMLFileList().count();
        for (int i = 0; i < sections; i++) {
            visitParaListCore(hwpx.sectionXMLFileList().get(i), from, to);
        }
    }

    private static void visitParaListCore(ParaListCore list,
                                          TablePageBreak from,
                                          TablePageBreak to) {
        int n = list.countOfPara();
        for (int i = 0; i < n; i++) visitPara(list.getPara(i), from, to);
    }

    private static void visitPara(Para para, TablePageBreak from, TablePageBreak to) {
        int rc = para.countOfRun();
        for (int i = 0; i < rc; i++) visitRun(para.getRun(i), from, to);
    }

    private static void visitRun(Run run, TablePageBreak from, TablePageBreak to) {
        int n = run.countOfRunItem();
        for (int i = 0; i < n; i++) {
            RunItem item = run.getRunItem(i);
            if (item instanceof Table) fixTable((Table) item, from, to);
        }
    }

    private static void fixTable(Table table, TablePageBreak from, TablePageBreak to) {
        if (table.pageBreak() == from) table.pageBreak(to);
        int rowCount = table.countOfTr();
        for (int r = 0; r < rowCount; r++) {
            Tr tr = table.getTr(r);
            int tcCount = tr.countOfTc();
            for (int c = 0; c < tcCount; c++) {
                Tc tc = tr.getTc(c);
                SubList sub = tc.subList();
                if (sub != null) visitParaListCore(sub, from, to);
            }
        }
    }
}
