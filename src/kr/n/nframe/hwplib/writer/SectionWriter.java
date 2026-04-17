package kr.n.nframe.hwplib.writer;

import kr.n.nframe.hwplib.binary.HwpBinaryWriter;
import kr.n.nframe.hwplib.binary.RecordWriter;
import kr.n.nframe.hwplib.constants.CtrlId;
import kr.n.nframe.hwplib.constants.HwpTagId;
import kr.n.nframe.hwplib.model.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * HWP 5.0 파일을 위한 BodyText/SectionN 스트림 레코드를 생성한다.
 * 원시(비압축) byte를 반환한다. 압축은 HwpWriter가 처리한다.
 *
 * 모든 레코드 포맷은 한글 워드프로세서가 생성한 참조 HWP 파일의 byte 분석과 정확히 일치한다.
 * 한글 워드프로세서가 생성한 결과.
 */
public class SectionWriter {

    private SectionWriter() {}

    /**
     * 단일 섹션의 모든 레코드를 생성한다.
     *
     * @param section 섹션 모델
     * @param doc     부모 문서 (필요시 컨텍스트 용도)
     * @return 원시 레코드 byte (비압축)
     */
    public static byte[] build(Section section, HwpDocument doc) {
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream(8192);
            for (int i = 0; i < section.paragraphs.size(); i++) {
                boolean lastInList = (i == section.paragraphs.size() - 1);
                writeParagraph(out, section.paragraphs.get(i), 0, false, lastInList);
            }
            return out.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException("Failed to build Section", e);
        }
    }

    // -----------------------------------------------------------------------
    // 문단 쓰기
    // -----------------------------------------------------------------------

    /**
     * 모든 하위 레코드를 포함한 완전한 문단을 작성한다.
     *
     * 레코드 순서:
     *   PARA_HEADER (level에 위치)
     *   PARA_TEXT (level+1에 위치) -- nChars > 1 또는 rawBytes > 2인 경우에만
     *   PARA_CHAR_SHAPE (level+1에 위치)
     *   PARA_LINE_SEG (level+1에 위치)
     *   컨트롤 (level+1에 위치)
     */
    private static void writeParagraph(ByteArrayOutputStream out, Paragraph para, int level,
                                       boolean inCell, boolean lastInList) throws IOException {
        // PARA_HEADER
        out.write(buildParaHeader(para, level, inCell, lastInList));

        // PARA_TEXT: 문단 끝 마커 이외의 실제 내용이 있는 경우에만 작성.
        // nChars=1이고 내용이 0x000D뿐이면 PARA_TEXT 자체를 생략한다.
        boolean hasText = para.paraText != null
                && para.paraText.rawBytes != null
                && para.paraText.rawBytes.length > 2; // 0x000D (2 byte) 이상일 경우

        if (hasText) {
            out.write(RecordWriter.buildRecord(HwpTagId.PARA_TEXT, level + 1, para.paraText.rawBytes));
        }

        // PARA_CHAR_SHAPE: 항목이 존재하면 작성 (빈 문단도 참조본에 있음)
        if (!para.charShapeRefs.isEmpty()) {
            out.write(buildParaCharShape(para.charShapeRefs, level + 1));
        }

        // PARA_LINE_SEG: 항목이 존재하면 작성
        if (!para.lineSegs.isEmpty()) {
            out.write(buildParaLineSeg(para.lineSegs, level + 1));
        }

        // PARA_RANGE_TAG: PARA_LINE_SEG 뒤에 기록 (한컴 참조 순서와 동일:
        //   HEADER, TEXT, CHAR_SHAPE, LINE_SEG, RANGE_TAG, 컨트롤...).
        // 하이퍼링크 텍스트가 한글에서 클릭 가능 / 스타일 적용으로 인식되기 위해 필요함.
        if (!para.rangeTags.isEmpty()) {
            out.write(buildParaRangeTag(para.rangeTags, level + 1));
        }

        // 컨트롤
        for (Control ctrl : para.controls) {
            writeControl(out, ctrl, level + 1);
        }
    }

    /**
     * PARA_HEADER: 22 byte.
     *
     * UINT32 nChars(4) + UINT32 controlMask(4) + UINT16 paraShapeId(2) +
     * UINT8 styleId(1) + UINT8 columnBreakType(1) + UINT16 charShapeCount(2) +
     * UINT16 rangeTagCount(2) + UINT16 lineSegCount(2) + UINT32 instanceId(4)
     *
     * nChars는 문단 텍스트의 총 WCHAR 개수:
     *   - 확장 컨트롤은 각각 8 WCHAR로 계산
     *   - 문자형 컨트롤(tab, line break, para end)은 각각 1 WCHAR로 계산
     *   - 빈 문단의 경우(PARA_TEXT 레코드 없음): nChars=1
     *
     * controlMask = 각 확장 컨트롤에 대한 (1 << charCode)의 OR 연산.
     *   - 문단 끝(0x0D)은 포함되지 않음
     *   - Tab(0x09)은 존재할 경우 포함됨
     *   - pghd (code 21) = bit 21 = 0x200000
     *   - gso (code 11) = bit 11 = 0x800
     *
     * instanceId: 대부분의 문단에 대해 0x80000000 (bit 31 설정).
     */
    private static byte[] buildParaHeader(Paragraph para, int level, boolean inCell, boolean lastInList) {
        HwpBinaryWriter w = new HwpBinaryWriter(28);

        boolean hasText = para.paraText != null
                && para.paraText.rawBytes != null
                && para.paraText.rawBytes.length > 2;

        long nChars = para.nChars;
        long controlMask = para.controlMask;
        long instanceId = para.instanceId & 0xFFFFFFFFL;

        if (!hasText) {
            // 빈 문단: nChars=1 (암시적 문단 개행), controlMask=0
            nChars = 1;
            controlMask = 0;
            instanceId = 0x80000000L;
        }

        // nChars bit 31 = "lastInList" 플래그.
        // hwplib는 셀/리스트에서 문단 읽기를 멈출 시점을 결정하기 위해 이 값을 사용한다.
        // 각 문단 리스트(섹션, 셀 등)의 마지막 문단에 반드시 설정되어야 한다.
        if (lastInList) {
            nChars |= 0x80000000L;
        }

        w.writeUInt32(nChars);                          // 4 byte
        w.writeUInt32(controlMask);                     // 4 byte
        w.writeUInt16(para.paraShapeId);                // 2 byte
        w.writeUInt8(para.styleId);                     // 1 byte
        w.writeUInt8(para.columnBreakType);             // 1 byte
        w.writeUInt16(para.charShapeRefs.size());       // 2 byte
        w.writeUInt16(para.rangeTags.size());           // 2 byte: rangeTagCount
        w.writeUInt16(para.lineSegs.size());            // 2 byte: lineSegCount
        w.writeUInt32(instanceId);                      // 4 byte
        w.writeUInt16(0);                               // 2 byte: mergeFlag
        return RecordWriter.buildRecord(HwpTagId.PARA_HEADER, level, w.toByteArray());
    }

    /**
     * PARA_CHAR_SHAPE: (UINT32 position + UINT32 charShapeId) 쌍의 배열.
     * 각 항목은 8 byte.
     */
    private static byte[] buildParaCharShape(List<CharShapeRef> refs, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(Math.multiplyExact(refs.size(), 8));
        for (CharShapeRef ref : refs) {
            w.writeUInt32(ref.position);
            w.writeUInt32(ref.charShapeId);
        }
        return RecordWriter.buildRecord(HwpTagId.PARA_CHAR_SHAPE, level, w.toByteArray());
    }

    /**
     * PARA_RANGE_TAG: 12 byte 항목의 배열.
     * 각 항목: UINT32 start + UINT32 end + UINT32 tag
     *   tag = (sort << 24) | (data & 0xFFFFFF)
     */
    private static byte[] buildParaRangeTag(List<ParaRangeTag> tags, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(Math.multiplyExact(tags.size(), 12));
        for (ParaRangeTag t : tags) {
            w.writeUInt32(t.start);
            w.writeUInt32(t.end);
            long tag = ((long)(t.sort & 0xFF) << 24) | (t.data & 0xFFFFFFL);
            w.writeUInt32(tag);
        }
        return RecordWriter.buildRecord(HwpTagId.PARA_RANGE_TAG, level, w.toByteArray());
    }

    /**
     * PARA_LINE_SEG: 36 byte 항목의 배열.
     * 각 항목: textStartPos(4) + lineVertPos(4) + lineHeight(4) + textHeight(4) +
     *       baselineDistance(4) + lineSpacing(4) + columnStartPos(4) + segWidth(4) + tag(4)
     */
    private static byte[] buildParaLineSeg(List<LineSeg> segs, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(Math.multiplyExact(segs.size(), 36));
        for (LineSeg seg : segs) {
            w.writeUInt32(seg.textStartPos);
            w.writeInt32(seg.lineVertPos);
            w.writeInt32(seg.lineHeight);
            w.writeInt32(seg.textHeight);
            w.writeInt32(seg.baselineDistance);
            w.writeInt32(seg.lineSpacing);
            w.writeInt32(seg.columnStartPos);
            w.writeInt32(seg.segWidth);
            w.writeUInt32(seg.tag);
        }
        return RecordWriter.buildRecord(HwpTagId.PARA_LINE_SEG, level, w.toByteArray());
    }

    // -----------------------------------------------------------------------
    // 컨트롤 쓰기
    // -----------------------------------------------------------------------

    private static void writeControl(ByteArrayOutputStream out, Control ctrl, int level)
            throws IOException {
        if (ctrl instanceof CtrlSectionDef) {
            writeSectionDef(out, (CtrlSectionDef) ctrl, level);
        } else if (ctrl instanceof CtrlColumnDef) {
            writeColumnDef(out, (CtrlColumnDef) ctrl, level);
        } else if (ctrl instanceof CtrlTable) {
            writeTable(out, (CtrlTable) ctrl, level);
        } else if (ctrl instanceof CtrlPicture) {
            writePicture(out, (CtrlPicture) ctrl, level);
        } else if (ctrl instanceof CtrlHeaderFooter) {
            writeHeaderFooter(out, (CtrlHeaderFooter) ctrl, level);
        } else if (ctrl instanceof CtrlField) {
            writeField(out, (CtrlField) ctrl, level);
        } else {
            writeGenericControl(out, ctrl, level);
        }
    }

    /**
     * 일반 컨트롤: ctrlId(4) + 원시 속성 데이터 byte.
     * pgnp (쪽 번호 위치), atno (자동 번호), nwno (새 번호),
     * pghd (쪽 숨기기), pgct (홀/짝), idxm (색인), bokm (책갈피) 등에 사용.
     */
    private static void writeGenericControl(ByteArrayOutputStream out, Control ctrl, int level)
            throws IOException {
        HwpBinaryWriter w = new HwpBinaryWriter(20);
        w.writeUInt32(ctrl.ctrlId);
        if (ctrl.rawData != null) {
            w.writeBytes(ctrl.rawData);
        } else {
            // 일부 컨트롤은 hwplib 리더가 기대하는 스펙으로 고정된 후속 데이터
            // 블록을 가진다. rawData가 없을 때(예: HWPX 소스에 hp:autoNumFormat
            // 세부 정보가 없는 경우)에도 레코드 길이가 스펙과 일치하도록
            // 0으로 채운 데이터를 반드시 기록해야 하며, 그렇지 않으면 한글은
            // 해당 레코드를 잘린 것으로 간주하여 이후 문단 레이아웃이 무너진다.
            int ctrlId = ctrl.ctrlId;
            if (ctrlId == CtrlId.AUTO_NUMBER || ctrlId == CtrlId.NEW_NUMBER) {
                // HWP 5.0 §4.3.10.5 자동 번호 / §4.3.10.6 새 번호 지정:
                //   UINT32 property(4) + UINT16 number(2)
                //   + WCHAR userChar(2) + WCHAR prefix(2) + WCHAR suffix(2)
                //   = 12 byte 데이터 → 총 16 byte 레코드.
                w.writePad(12);
            } else if (ctrlId == CtrlId.PAGE_HIDE || ctrlId == CtrlId.PAGE_ODD_EVEN) {
                // §4.3.10.7~8 — UINT32 속성 = 4 byte 데이터.
                w.writePad(4);
            } else if (ctrlId == CtrlId.PAGE_NUM_POS) {
                // §4.3.10.9 — UINT32 속성 + UINT16 number_position = 6 byte.
                w.writePad(6);
            } else if (ctrlId == CtrlId.INDEX_MARK) {
                // §4.3.10.10 — 최소 6 byte (빈 키워드).
                w.writePad(6);
            }
            // BOOKMARK (bokm)는 HWPTAG_CTRL_DATA 형제 레코드를 통해 데이터를
            // 전달하므로 여기에는 후속 byte가 없다.
        }
        out.write(RecordWriter.buildRecord(HwpTagId.CTRL_HEADER, level, w.toByteArray()));
    }

    // --- CtrlSectionDef ---

    /**
     * 섹션 정의용 CTRL_HEADER: 총 47 byte.
     *
     * ctrlId(4) + property(4) + columnGap(2) + vertGrid(2) + horizGrid(2) +
     * defaultTabStop(4) + numberingParaShapeId(2) + pageNumber(2) + figNumber(2) +
     * tableNumber(2) + equationNumber(2) + defaultLang(2) = 30 byte
     * + 버전 5.0.1.5+ 호환성을 위한 17 byte 0 패딩 = 47 byte
     *
     * 참조본 byte 18 (ctrlId 시작부터의 offset) = outlineShapeIDRef (UINT16), 값=1.
     * 이는 numberingParaShapeId 위치에 해당한다.
     */
    private static void writeSectionDef(ByteArrayOutputStream out, CtrlSectionDef sd, int level)
            throws IOException {
        HwpBinaryWriter w = new HwpBinaryWriter(48);
        w.writeUInt32(sd.ctrlId);              // 4 byte
        w.writeUInt32(sd.property);            // 4 byte
        w.writeHwpUnit16(sd.columnGap);        // 2 byte
        w.writeHwpUnit16(sd.vertGrid);         // 2 byte
        w.writeHwpUnit16(sd.horizGrid);        // 2 byte
        w.writeHwpUnit(sd.defaultTabStop);     // 4 byte
        w.writeUInt16(sd.numberingParaShapeId); // 2 byte
        w.writeUInt16(sd.pageNumber);          // 2 byte
        w.writeUInt16(sd.figNumber);           // 2 byte
        w.writeUInt16(sd.tableNumber);         // 2 byte
        w.writeUInt16(sd.equationNumber);      // 2 byte
        w.writeUInt16(sd.defaultLang);         // 2 byte
        // --- 여기까지 30 byte ---
        w.writePad(17);                        // 17 byte 0 패딩
        // --- 총 47 byte ---
        out.write(RecordWriter.buildRecord(HwpTagId.CTRL_HEADER, level, w.toByteArray()));

        // CTRL_DATA는 여기에 작성하지 않음 - 일부 참조 파일에만 존재하며
        // 대부분의 파일에서는 없어도 문제가 되지 않는다.

        // PAGE_DEF
        out.write(buildPageDef(sd.pageDef, level + 1));

        // FOOTNOTE_SHAPE (각주)
        out.write(buildFootnoteShape(sd.footnoteShape, level + 1));

        // FOOTNOTE_SHAPE (미주)
        out.write(buildFootnoteShape(sd.endnoteShape, level + 1));

        // PAGE_BORDER_FILL 레코드
        for (PageBorderFill pbf : sd.pageBorderFills) {
            out.write(buildPageBorderFill(pbf, level + 1));
        }
    }

    /**
     * PAGE_DEF: 40 byte.
     * HWPUNIT paperWidth(4) + paperHeight(4) + leftMargin(4) + rightMargin(4) +
     * topMargin(4) + bottomMargin(4) + headerMargin(4) + footerMargin(4) +
     * gutterMargin(4) + UINT32 property(4)
     */
    private static byte[] buildPageDef(PageDef pd, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(40);
        w.writeHwpUnit(pd.paperWidth);
        w.writeHwpUnit(pd.paperHeight);
        w.writeHwpUnit(pd.leftMargin);
        w.writeHwpUnit(pd.rightMargin);
        w.writeHwpUnit(pd.topMargin);
        w.writeHwpUnit(pd.bottomMargin);
        w.writeHwpUnit(pd.headerMargin);
        w.writeHwpUnit(pd.footerMargin);
        w.writeHwpUnit(pd.gutterMargin);
        w.writeUInt32(pd.property);
        return RecordWriter.buildRecord(HwpTagId.PAGE_DEF, level, w.toByteArray());
    }

    /**
     * FOOTNOTE_SHAPE: 정확히 28 byte.
     *
     * UINT32 property(4) + WCHAR userSymbol(2) + WCHAR prefixChar(2) +
     * WCHAR suffixChar(2) + UINT16 startNumber(2) + INT32 dividerLength(4) +
     * HWPUNIT16 aboveMargin(2) + HWPUNIT16 belowMargin(2) +
     * HWPUNIT16 betweenNote(2) + UINT8 dividerLineType(1) +
     * UINT8 dividerLineWidth(1) + COLORREF dividerColor(4) = 28 byte
     *
     * dividerLength는 INT32 (4 byte)이며 INT16이 아님! 자동의 경우 -1 = 0xFFFFFFFF.
     * dividerLineType: 0=없음, 1=실선 (테두리 선 타입 enum과 다름).
     * dividerLineWidth: 0=0.1mm, 1=0.12mm 등.
     */
    private static byte[] buildFootnoteShape(FootnoteShape fs, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(32);
        w.writeUInt32(fs.property);              // 4 byte
        w.writeWChar(fs.userSymbol);             // 2 byte
        w.writeWChar(fs.prefixChar);             // 2 byte
        w.writeWChar(fs.suffixChar);             // 2 byte
        w.writeUInt16(fs.startNumber);           // 2 byte
        w.writeInt32(fs.dividerLength);          // 4 byte (INT32!)
        w.writeHwpUnit16(fs.dividerAboveMargin); // 2 byte
        w.writeHwpUnit16(fs.dividerBelowMargin); // 2 byte
        w.writeHwpUnit16(fs.noteBetweenMargin);  // 2 byte
        w.writeUInt8(fs.dividerLineType);        // 1 byte
        w.writeUInt8(fs.dividerLineWidth);       // 1 byte
        w.writeColorRef(fs.dividerLineColor);    // 4 byte
        // --- 총 28 byte ---
        return RecordWriter.buildRecord(HwpTagId.FOOTNOTE_SHAPE, level, w.toByteArray());
    }

    /**
     * PAGE_BORDER_FILL: 14 byte.
     * UINT32 property(4) + HWPUNIT16 padL(2) + padR(2) + padT(2) + padB(2) +
     * UINT16 borderFillId(2)
     *
     * Property 비트:
     *   bit 0:    위치 기준 (0=본문, 1=용지)
     *   bit 1:    머리말 포함
     *   bit 2:    꼬리말 포함
     *   bits 3-4: 채움 영역 (0=용지, 1=페이지, 2=테두리)
     */
    private static byte[] buildPageBorderFill(PageBorderFill pbf, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(16);
        w.writeUInt32(pbf.property);
        w.writeHwpUnit16(pbf.padLeft);
        w.writeHwpUnit16(pbf.padRight);
        w.writeHwpUnit16(pbf.padTop);
        w.writeHwpUnit16(pbf.padBottom);
        w.writeUInt16(pbf.borderFillId);
        return RecordWriter.buildRecord(HwpTagId.PAGE_BORDER_FILL, level, w.toByteArray());
    }

    // --- CtrlColumnDef ---

    /**
     * 단 정의 컨트롤.
     * ctrlId(4) + propertyLow(2) + columnGap(2) + [columnWidths] +
     * propertyHigh(2) + dividerType(1) + dividerWidth(1) + dividerColor(4)
     */
    private static void writeColumnDef(ByteArrayOutputStream out, CtrlColumnDef cd, int level)
            throws IOException {
        HwpBinaryWriter w = new HwpBinaryWriter(64);
        w.writeUInt32(cd.ctrlId);
        w.writeUInt16(cd.propertyLow);
        w.writeHwpUnit16(cd.columnGap);

        // 같은 크기가 아닌 경우의 컬럼 너비
        if (cd.columnWidths != null && cd.columnWidths.length > 0) {
            for (int cw : cd.columnWidths) {
                w.writeUInt16(cw);
            }
        }

        w.writeUInt16(cd.propertyHigh);
        w.writeUInt8(cd.dividerType);
        w.writeUInt8(cd.dividerWidth);
        w.writeColorRef(cd.dividerColor);
        out.write(RecordWriter.buildRecord(HwpTagId.CTRL_HEADER, level, w.toByteArray()));
    }

    // --- CtrlTable ---

    /**
     * 표 컨트롤 쓰기.
     * 공통 개체 속성을 갖는 CTRL_HEADER 다음에 TABLE 레코드, 그리고 셀 순서.
     *
     * TABLE property 참조 값: 0x00000006 = bits 0-1=2 (셀 단위 분할) + bit 2=1 (머리 반복).
     * "CELL" pageBreak는 bits 0-1에서 값 2에 해당 (1 아님!).
     */
    private static void writeTable(ByteArrayOutputStream out, CtrlTable tbl, int level)
            throws IOException {
        // 공통 개체 속성을 갖는 CTRL_HEADER
        out.write(buildCommonCtrlHeader(tbl.ctrlId, tbl.objProperty, tbl.vertOffset,
                tbl.horzOffset, tbl.width, tbl.height, tbl.zOrder, tbl.margins,
                tbl.instanceId, tbl.pageBreakPrev, tbl.description, level));

        // TABLE 레코드
        out.write(buildTableRecord(tbl, level + 1));

        // 셀
        for (TableCell cell : tbl.cells) {
            writeTableCell(out, cell, level + 1);
        }
    }

    /**
     * TABLE 레코드: rowCount에 따라 가변 크기.
     *
     * UINT32 property(4) + UINT16 rowCount(2) + UINT16 colCount(2) +
     * HWPUNIT16 cellSpacing(2) + HWPUNIT16 padL(2) + padR(2) + padT(2) + padB(2) +
     * WORD[rowCount] rowHeights(2*rowCount) + UINT16 borderFillId(2) +
     * UINT16 zoneCount(2)
     *
     * 참조본: 28 = 4+2+2+2+8+6+2+2 = 28 (3행: rowHeights 6 byte)
     * 참조본의 행 높이 값: 1,1,1 (placeholder 값).
     */
    private static byte[] buildTableRecord(CtrlTable tbl, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(64);
        w.writeUInt32(tbl.property);            // 4 byte
        w.writeUInt16(tbl.rowCount);            // 2 byte
        w.writeUInt16(tbl.colCount);            // 2 byte
        w.writeHwpUnit16(tbl.cellSpacing);      // 2 byte
        w.writeHwpUnit16(tbl.padLeft);          // 2 byte
        w.writeHwpUnit16(tbl.padRight);         // 2 byte
        w.writeHwpUnit16(tbl.padTop);           // 2 byte
        w.writeHwpUnit16(tbl.padBottom);        // 2 byte

        // 행별 높이 (행당 WORD)
        if (tbl.rowHeights != null) {
            for (int rh : tbl.rowHeights) {
                w.writeUInt16(rh);
            }
        }

        w.writeUInt16(tbl.borderFillId);        // 2 byte
        w.writeUInt16(tbl.zoneCount);           // 2 byte
        // zone 데이터 생략 (zoneCount는 보통 0)

        return RecordWriter.buildRecord(HwpTagId.TABLE, level, w.toByteArray());
    }

    /**
     * 표 셀용 LIST_HEADER: 정확히 47 byte.
     *
     * 리스트 헤더 기본(8 byte):
     *   INT16 paraCount(2) + UINT32 listProperty(4) + UINT16 padding(2)
     *
     * 셀 전용 속성(26 byte):
     *   UINT16 colAddr(2) + UINT16 rowAddr(2) + UINT16 colSpan(2) + UINT16 rowSpan(2) +
     *   HWPUNIT width(4) + HWPUNIT height(4) + HWPUNIT16[4] margins(8) +
     *   UINT16 borderFillId(2)
     *
     * 추가 버전 필드(13 byte):
     *   HWPUNIT width_copy(4) + BYTE[9] padding(9)
     *
     * 총합: 8 + 26 + 13 = 47 byte
     *
     * listProperty: vertAlign이 bit 21에 위치 (bit 5 아님!)
     *
     * 셀 문단은 LIST_HEADER와 동일한 level에 작성됨 (level+1 아님).
     */
    private static void writeTableCell(ByteArrayOutputStream out, TableCell cell, int level)
            throws IOException {
        HwpBinaryWriter w = new HwpBinaryWriter(48);

        // hwplib 기준 list header base (8 byte): readSInt4 + readUInt4
        w.writeInt32(cell.paragraphs.size());    // 4 byte: paraCount (INT32!)
        w.writeUInt32(cell.listHeaderProperty);  // 4 byte: vertAlign이 bit 21에 위치

        // 셀 전용 속성: colIndex, rowIndex, colSpan, rowSpan, width, height...
        w.writeUInt16(cell.colAddr);             // 2 byte: colIndex
        w.writeUInt16(cell.rowAddr);             // 2 byte: rowIndex
        w.writeUInt16(cell.colSpan);             // 2 byte
        w.writeUInt16(cell.rowSpan);             // 2 byte
        w.writeHwpUnit(cell.width);              // 4 byte
        w.writeHwpUnit(cell.height);             // 4 byte
        for (int i = 0; i < 4; i++) {
            w.writeHwpUnit16(cell.margins[i]);   // 총 8 byte (L,R,T,B)
        }
        w.writeUInt16(cell.borderFillId);        // 2 byte

        // 추가 버전 필드 (13 byte)
        w.writeHwpUnit(cell.width);              // 4 byte: width 복사본
        w.writePad(9);                           // 9 byte: 0 패딩
        // 총합: 8 + 26 + 13 = 47 byte

        out.write(RecordWriter.buildRecord(HwpTagId.LIST_HEADER, level, w.toByteArray()));

        // 셀 문단은 LIST_HEADER와 동일한 level에 위치
        // 마지막 문단은 nChars bit 31이 설정되어야 함 (lastInList 플래그)
        for (int pi = 0; pi < cell.paragraphs.size(); pi++) {
            boolean isLast = (pi == cell.paragraphs.size() - 1);
            writeParagraph(out, cell.paragraphs.get(pi), level, true, isLast);
        }
    }

    // --- CtrlPicture ---

    /**
     * 그림 컨트롤 쓰기.
     *
     * 기록되는 레코드:
     *   CTRL_HEADER (level): ctrlId('gso ') + 공통 개체 속성 (46+2*descLen byte)
     *   SHAPE_COMPONENT (level+1): ctrlId 두 번 기록 + 변환 + 렌더링 행렬
     *   SHAPE_COMPONENT_PICTURE (level+2): 테두리 + imageRect + crop + padding + imageInfo
     *
     * gso CTRL_HEADER 크기 = 46 + 2*descLen. 72자 description의 경우 = 190 byte.
     * description 문자열은 HWPX picture 요소로부터 정확히 전달되어야 한다.
     */
    private static void writePicture(ByteArrayOutputStream out, CtrlPicture pic, int level)
            throws IOException {
        // 공통 개체 속성을 갖는 CTRL_HEADER (ctrlId는 GSO, PICTURE 아님)
        out.write(buildCommonCtrlHeader(pic.ctrlId, pic.objProperty, pic.vertOffset,
                pic.horzOffset, pic.width, pic.height, pic.zOrder, pic.margins,
                pic.instanceId, pic.pageBreakPrev, pic.description, level));

        // 수식 개체(ctrlId "eqed")는 CTRL_HEADER 뒤에 shape-component 레코드
        // 대신 HWPTAG_EQEDIT 페이로드를 담는다. 이것이 없으면 한글에서 수식
        // 영역이 공백으로 표시되며 — 이는 TA-04의 수식 셀에서 관측된 버그이다.
        if ("equation".equals(pic.shapeType)) {
            out.write(buildEquationRecord(pic, level + 1));
            return;
        }

        // shapeType 기반으로 shape별 ctrlId 결정
        int shapeCtrlId = CtrlId.PICTURE; // 'pic'에 대한 기본값
        int shapeTag = HwpTagId.SHAPE_COMPONENT_PICTURE; // shape 레코드의 기본 태그
        if (pic.shapeType != null) {
            switch (pic.shapeType) {
                case "rect": shapeCtrlId = CtrlId.RECTANGLE; shapeTag = HwpTagId.SHAPE_COMPONENT_RECT; break;
                case "line": shapeCtrlId = CtrlId.LINE; shapeTag = HwpTagId.SHAPE_COMPONENT_LINE; break;
                case "ellipse": shapeCtrlId = CtrlId.ELLIPSE; shapeTag = HwpTagId.SHAPE_COMPONENT_ELLIPSE; break;
                case "arc": shapeCtrlId = CtrlId.ARC; shapeTag = HwpTagId.SHAPE_COMPONENT_ARC; break;
                case "polygon": shapeCtrlId = CtrlId.POLYGON; shapeTag = HwpTagId.SHAPE_COMPONENT_POLYGON; break;
                case "curve": shapeCtrlId = CtrlId.CURVE; shapeTag = HwpTagId.SHAPE_COMPONENT_CURVE; break;
                default: shapeCtrlId = CtrlId.PICTURE; break;
            }
        }

        // SHAPE_COMPONENT (tag 0x04C)는 level+1에 위치
        // GenShapeObject의 경우 ctrlId가 두 번 기록된다.
        HwpBinaryWriter sw = new HwpBinaryWriter(256);
        sw.writeUInt32(shapeCtrlId);       // 4: shape 전용 ctrlId (1회차 기록)
        sw.writeUInt32(shapeCtrlId);       // 4: shape 전용 ctrlId (2회차 기록)

        // initialWidth/Height = 원본 이미지 크기 (스케일링 이전)
        // currentWidth/Height = 표시 크기 (스케일링 이후)
        int initW = pic.originalWidth > 0 ? pic.originalWidth : pic.width;
        int initH = pic.originalHeight > 0 ? pic.originalHeight : pic.height;
        int curW = pic.width;
        int curH = pic.height;

        // xOffset/yOffset: HWPX transMatrix e3/e6에서 가져오거나 계산한 값
        int scXOff, scYOff;
        if (pic.transMatrix != null) {
            scXOff = (int) pic.transMatrix[2]; // e3 = tx
            scYOff = (int) pic.transMatrix[5]; // e6 = ty
        } else {
            scXOff = -(initW - curW) / 2;
            scYOff = -(initH - curH) / 2;
        }

        sw.writeInt32(scXOff);             // 4: 그룹 내 xOffset
        sw.writeInt32(scYOff);             // 4: 그룹 내 yOffset
        sw.writeUInt16(0);                 // 2: groupLevel
        sw.writeUInt16(1);                 // 2: localVersion
        sw.writeHwpUnit(initW);            // 4: initialWidth (원본 크기)
        sw.writeHwpUnit(initH);            // 4: initialHeight (원본 크기)
        sw.writeHwpUnit(curW);             // 4: currentWidth (표시 크기)
        sw.writeHwpUnit(curH);             // 4: currentHeight (표시 크기)
        // property: 스케일링 된 그림은 bit 19, 일부 플래그는 bit 29
        long scProp = 0;
        if (initW != curW || initH != curH) {
            scProp = 0x24080000L; // 스케일링 된 그림의 참조본에서 관측됨
        }
        sw.writeUInt32(scProp);            // 4: property (flip/scale 플래그)
        sw.writeUInt16(0);                 // 2: 회전 각도 (도)
        sw.writeInt32(curW / 2);           // 4: 회전 중심 X
        sw.writeInt32(curH / 2);           // 4: 회전 중심 Y

        // 렌더링 정보: WORD matrixPairCount + 행렬들
        // 각 행렬은 3x2 아핀: [e1, e2, e3, e4, e5, e6], 6개 double로 저장 (48 byte)
        sw.writeUInt16(1);                 // 2: 행렬 쌍 개수 = 1

        if (pic.transMatrix != null && pic.scaMatrix != null && pic.rotMatrix != null) {
            // HWPX renderingInfo의 정확한 값 사용
            for (double v : pic.transMatrix) writeDouble(sw, v);
            for (double v : pic.scaMatrix)   writeDouble(sw, v);
            for (double v : pic.rotMatrix)   writeDouble(sw, v);
        } else {
            // 크기로부터 행렬 계산
            double scaleX = (initW > 0) ? (double) curW / initW : 1.0;
            double scaleY = (initH > 0) ? (double) curH / initH : 1.0;

            // 이동: [1, 0, xOff, 0, 1, yOff]
            writeDouble(sw, 1.0); writeDouble(sw, 0.0); writeDouble(sw, (double) scXOff);
            writeDouble(sw, 0.0); writeDouble(sw, 1.0); writeDouble(sw, (double) scYOff);

            // 스케일: [scaleX, 0, -xOff, 0, scaleY, -yOff]
            writeDouble(sw, scaleX); writeDouble(sw, 0.0); writeDouble(sw, (double) -scXOff);
            writeDouble(sw, 0.0); writeDouble(sw, scaleY); writeDouble(sw, (double) -scYOff);

            // 회전: 단위 행렬
            writeDouble(sw, 1.0); writeDouble(sw, 0.0); writeDouble(sw, 0.0);
            writeDouble(sw, 0.0); writeDouble(sw, 1.0); writeDouble(sw, 0.0);
        }

        out.write(RecordWriter.buildRecord(HwpTagId.SHAPE_COMPONENT, level + 1, sw.toByteArray()));

        // 텍스트박스를 가진 그리기 개체: LIST_HEADER + 문단이 shape 전용 레코드보다 먼저 옴
        // 참조 순서: SHAPE_COMPONENT -> LIST_HEADER -> 문단 -> SHAPE_COMPONENT_XXX
        if (!pic.textboxParagraphs.isEmpty()) {
            // 텍스트박스 LIST_HEADER 형식 (hwplib ForTextBox 리더와 일치):
            //   INT32 paraCount(4) + UINT32 property(4) +
            //   UINT16 margins[4](8) + UINT32 textWidth(4)
            //   = 기본 20 byte, 13 byte 버전 패딩 추가 = 총 33 byte
            HwpBinaryWriter lw = new HwpBinaryWriter(48);
            lw.writeInt32(pic.textboxParagraphs.size());    // 4 byte: paraCount (INT32, INT16 아님)
            lw.writeUInt32(pic.textboxListProperty);        // 4 byte: 리스트 속성
            lw.writeHwpUnit16(pic.textboxMarginLeft);       // 2 byte
            lw.writeHwpUnit16(pic.textboxMarginRight);      // 2 byte
            lw.writeHwpUnit16(pic.textboxMarginTop);        // 2 byte
            lw.writeHwpUnit16(pic.textboxMarginBottom);     // 2 byte
            lw.writeHwpUnit(pic.textboxLastWidth);          // 4 byte: textWidth
            lw.writePad(13);                                // 13 byte: 버전 패딩
            // 총합: 4+4+8+4+13 = 33 byte
            out.write(RecordWriter.buildRecord(HwpTagId.LIST_HEADER, level + 2, lw.toByteArray()));

            // 문단은 LIST_HEADER와 동일한 level(level+2)에 위치
            for (int pi = 0; pi < pic.textboxParagraphs.size(); pi++) {
                boolean isLast = (pi == pic.textboxParagraphs.size() - 1);
                writeParagraph(out, pic.textboxParagraphs.get(pi), level + 2, false, isLast);
            }
        }

        // Shape 전용 레코드는 level+2에 위치
        if ("pic".equals(pic.shapeType) || pic.shapeType == null) {
            // SHAPE_COMPONENT_PICTURE (tag 0x055)
            HwpBinaryWriter pw = new HwpBinaryWriter(128);
            pw.writeColorRef(pic.borderColor);
            pw.writeInt32(pic.borderThickness);
            pw.writeUInt32(pic.borderProperty);
            for (int i = 0; i < 4; i++) pw.writeInt32(pic.imageRectX[i]);
            for (int i = 0; i < 4; i++) pw.writeInt32(pic.imageRectY[i]);
            pw.writeInt32(pic.cropLeft);
            pw.writeInt32(pic.cropTop);
            pw.writeInt32(pic.cropRight);
            pw.writeInt32(pic.cropBottom);
            pw.writeHwpUnit16(pic.imgPadLeft);
            pw.writeHwpUnit16(pic.imgPadRight);
            pw.writeHwpUnit16(pic.imgPadTop);
            pw.writeHwpUnit16(pic.imgPadBottom);
            pw.writeUInt8(pic.brightness);
            pw.writeUInt8(pic.contrast);
            pw.writeUInt8(pic.effect);
            pw.writeUInt16(pic.binItemId);
            pw.writeUInt8(pic.transparency);
            pw.writeUInt32(pic.pictureInstanceId);
            pw.writeUInt32(0); // 그림 효과 정보
            // SHAPE_COMPONENT_PICTURE는 orgSz가 아닌 imgDim(원본 픽셀 크기)을 사용
            int picOrigW = pic.imgDimWidth > 0 ? pic.imgDimWidth : pic.originalWidth;
            int picOrigH = pic.imgDimHeight > 0 ? pic.imgDimHeight : pic.originalHeight;
            pw.writeHwpUnit(picOrigW);
            pw.writeHwpUnit(picOrigH);
            pw.writeUInt8(0); // 이미지 투명도
            out.write(RecordWriter.buildRecord(HwpTagId.SHAPE_COMPONENT_PICTURE, level + 2, pw.toByteArray()));
        } else if ("rect".equals(pic.shapeType)) {
            // SHAPE_COMPONENT_RECTANGLE: cornerRadius(1) + 4 x좌표(16) + 4 y좌표(16) = 33 byte
            HwpBinaryWriter pw = new HwpBinaryWriter(64);
            pw.writeUInt8(0); // 모서리 반경 %
            pw.writeInt32(0); pw.writeInt32(pic.width); pw.writeInt32(pic.width); pw.writeInt32(0); // X 좌표
            pw.writeInt32(0); pw.writeInt32(0); pw.writeInt32(pic.height); pw.writeInt32(pic.height); // Y 좌표
            out.write(RecordWriter.buildRecord(shapeTag, level + 2, pw.toByteArray()));
        } else if ("line".equals(pic.shapeType)) {
            // SHAPE_COMPONENT_LINE: startX(4)+startY(4)+endX(4)+endY(4)+flags(2) = 18 byte
            HwpBinaryWriter pw = new HwpBinaryWriter(32);
            pw.writeInt32(0); pw.writeInt32(0);
            pw.writeInt32(pic.width); pw.writeInt32(pic.height);
            pw.writeUInt16(0);
            out.write(RecordWriter.buildRecord(shapeTag, level + 2, pw.toByteArray()));
        } else if ("ellipse".equals(pic.shapeType)) {
            // SHAPE_COMPONENT_ELLIPSE: 60 byte (간략화)
            HwpBinaryWriter pw = new HwpBinaryWriter(64);
            pw.writeUInt32(0); // 속성
            int cx = pic.width / 2, cy = pic.height / 2;
            pw.writeInt32(cx); pw.writeInt32(cy); // 중심
            pw.writeInt32(pic.width); pw.writeInt32(0); // 축1
            pw.writeInt32(0); pw.writeInt32(pic.height); // 축2
            pw.writePad(32); // 시작/끝 위치
            out.write(RecordWriter.buildRecord(shapeTag, level + 2, pw.toByteArray()));
        } else {
            // 일반: 빈 shape 레코드 기록
            out.write(RecordWriter.buildRecord(shapeTag, level + 2, new byte[0]));
        }
    }

    // --- CtrlHeaderFooter ---

    /**
     * 머리말/꼬리말 컨트롤.
     * CTRL_HEADER: ctrlId(4) + property(4) = 8 byte
     * LIST_HEADER: paraCount(2) + listProp(4) + textWidth(4) + textHeight(4) +
     *              textRefFlag(1) + numRefFlag(1) = 16 byte
     * 그 다음 문단이 level+2에 위치.
     */
    private static void writeHeaderFooter(ByteArrayOutputStream out, CtrlHeaderFooter hf, int level)
            throws IOException {
        // CTRL_HEADER
        // HEADER/FOOTER CTRL_HEADER — 한컴 참조본과 일치시키기 위해 12 byte:
        //   ctrlId(4) + property(4) + 예약된 0 4 byte.
        // (한글이 생성하는 HWP는 일관되게 후속 4 byte를 포함하며,
        //  엄격한 리더가 추가 슬롯을 사용하므로 반드시 기록해야 한다.)
        HwpBinaryWriter w = new HwpBinaryWriter(32);
        w.writeUInt32(hf.ctrlId);
        w.writeUInt32(hf.property);
        w.writeUInt32(0);
        out.write(RecordWriter.buildRecord(HwpTagId.CTRL_HEADER, level, w.toByteArray()));

        // 머리말/꼬리말용 LIST_HEADER — 한컴 참조본과 일치시키기 위해 34 byte.
        // 스펙(표 140 머리말/꼬리말)에는 14 byte의 데이터가 나열되지만,
        // 한글은 추가 예약/페이지레이아웃 슬롯을 포함한 34 byte 블록으로
        // 감싼다. 우리가 검사한 모든 한글 생성 HWP에서 추가 18 byte는
        // 0으로 끝난다.
        HwpBinaryWriter lw = new HwpBinaryWriter(40);
        lw.writeInt16(hf.paragraphs.size());
        lw.writeUInt32(0);                 // listProperty (스펙 "속성")
        lw.writeHwpUnit(hf.textWidth);     // 텍스트 폭
        lw.writeHwpUnit(hf.textHeight);    // 텍스트 높이
        lw.writeUInt8(hf.textRefFlag);
        lw.writeUInt8(hf.numRefFlag);
        lw.writePad(18);                   // 예약된 후속 byte
        out.write(RecordWriter.buildRecord(HwpTagId.LIST_HEADER, level + 1, lw.toByteArray()));

        // 한컴 참조본에서 문단은 LIST_HEADER의 동급 형제(level+1)이며 —
        // 예전 코드의 level+2가 아니다. 올바른 깊이에 기록하면 적절한
        // 부모/자식 관계가 유지되어 한글이 다음 섹션의 페이지 레이아웃을
        // 정확히 계산한다 (표-텍스트 간 간격 붕괴 수정).
        for (int pi = 0; pi < hf.paragraphs.size(); pi++) {
            boolean isLast = (pi == hf.paragraphs.size() - 1);
            writeParagraph(out, hf.paragraphs.get(pi), level + 1, false, isLast);
        }
    }

    // --- CtrlField ---

    /**
     * 필드 컨트롤 (하이퍼링크, 책갈피, clickhere, 수식, 날짜 등).
     * ctrlId(4) + fieldProperty(4) + extraProperty(1) +
     * WORD cmdLen + WCHAR[] command + UINT32 fieldId
     */
    private static void writeField(ByteArrayOutputStream out, CtrlField fld, int level)
            throws IOException {
        HwpBinaryWriter w = new HwpBinaryWriter(64);
        w.writeUInt32(fld.ctrlId);
        w.writeUInt32(fld.fieldProperty);
        w.writeUInt8(fld.extraProperty);
        // 명령 문자열
        String cmd = (fld.command != null) ? fld.command : "";
        w.writeUInt16(cmd.length());
        w.writeUtf16Le(cmd);
        w.writeUInt32(fld.fieldId);
        w.writeUInt32(0); // 추가 4 byte 패딩 (참조본에서 관측됨)
        out.write(RecordWriter.buildRecord(HwpTagId.CTRL_HEADER, level, w.toByteArray()));

        // BOOKMARK 필드용 HWPTAG_CTRL_DATA: 책갈피 이름을 ParameterSet 형태로
        // 전달한다. 문서를 열 때 "?section_A"와 같은 내부 하이퍼링크 명령이
        // 올바른 앵커로 해석되도록 하기 위해 필요하다.
        // 스펙: HWP 5.0 §4.3.10.11 (책갈피) + §4.2.12 표 50 (Parameter Set).
        if (fld.ctrlId == CtrlId.FIELD_BOOKMARK && fld.name != null && !fld.name.isEmpty()) {
            out.write(buildBookmarkCtrlData(fld.name, level + 1));
        }
    }

    /**
     * 수식 개체용 HWPTAG_EQEDIT 레코드 생성. byte 레이아웃
     * (한글 참조본 TA-04에서 관측되었고 HWP 5.0 수식 스펙과 일치):
     * <pre>
     *   UINT32 property           (보통 0)
     *   WORD   scriptLen          (WCHAR 개수)
     *   WCHAR  script[]           수식 텍스트 (예: "15,000 × 240 = ?")
     *   UINT32 baseUnit           HWPUNIT (기본 1000 = 1 문자)
     *   UINT32 reserved           0
     *   UINT32 baseLine           텍스트 기준 백분율 (기본 86)
     *   WORD   versionLen
     *   WCHAR  version[]          "Equation Version 60"
     *   WORD   fontLen
     *   WCHAR  font[]             "HancomEQN"
     * </pre>
     */
    private static byte[] buildEquationRecord(CtrlPicture eq, int level) {
        String script = eq.eqScript != null ? eq.eqScript : "";
        String version = eq.eqVersion != null ? eq.eqVersion : "Equation Version 60";
        String font = eq.eqFont != null ? eq.eqFont : "HancomEQN";

        HwpBinaryWriter w = new HwpBinaryWriter(64 + 2 * (script.length() + version.length() + font.length()));
        w.writeUInt32(eq.eqProperty);
        w.writeUInt16(script.length());
        w.writeUtf16Le(script);
        w.writeUInt32(eq.eqBaseUnit);
        w.writeUInt32(0);                 // 예약
        w.writeUInt32(eq.eqBaseLine);
        w.writeUInt16(version.length());
        w.writeUtf16Le(version);
        w.writeUInt16(font.length());
        w.writeUtf16Le(font);
        return RecordWriter.buildRecord(HwpTagId.EQEDIT, level, w.toByteArray());
    }

    /**
     * 책갈피 이름용 HWPTAG_CTRL_DATA 레코드 생성.
     *
     * byte 레이아웃 (한컴 참조 출력과 일치):
     *   WORD   setID          = 0x021B
     *   INT16  itemCount      = 1
     *   UINT32 itemHeader     = 0x40000000  (item id = 0, high flag = 0x4000)
     *   WORD   itemType       = 1  (PIT_BSTR)
     *   WORD   strLen         = N  (WCHAR 개수)
     *   WCHAR  name[strLen]   (UTF-16 LE)
     * 총합: 12 + 2*N byte.
     */
    private static byte[] buildBookmarkCtrlData(String name, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(12 + 2 * name.length());
        w.writeUInt16(0x021B);     // paramSetID
        w.writeInt16(1);           // 항목 개수
        w.writeUInt32(0x40000000L);// 항목 헤더 (id=0, flag=0x4000)
        w.writeUInt16(1);          // itemType = PIT_BSTR
        w.writeUInt16(name.length());
        w.writeUtf16Le(name);
        return RecordWriter.buildRecord(HwpTagId.CTRL_DATA, level, w.toByteArray());
    }

    // -----------------------------------------------------------------------
    // 공용 헬퍼
    // -----------------------------------------------------------------------

    /**
     * 공통 개체 속성을 포함한 CTRL_HEADER 레코드 생성.
     * CtrlTable과 CtrlPicture 양쪽에서 사용.
     *
     * 형식: ctrlId(4) + objProperty(4) + vertOffset(4) + horzOffset(4) +
     *       width(4) + height(4) + zOrder(4) + margins[4](8) + instanceId(4) +
     *       pageBreakPrev(4) + descLen(2) + desc(2*descLen)
     * 총합: 46 + 2*descLen byte
     */
    private static byte[] buildCommonCtrlHeader(int ctrlId, long objProperty,
            int vertOffset, int horzOffset, int width, int height,
            int zOrder, int[] margins, long instanceId, int pageBreakPrev,
            String description, int level) {
        HwpBinaryWriter w = new HwpBinaryWriter(256);
        w.writeUInt32(ctrlId);             // 4 byte
        w.writeUInt32(objProperty);        // 4 byte
        w.writeHwpUnit(vertOffset);        // 4 byte
        w.writeHwpUnit(horzOffset);        // 4 byte
        w.writeHwpUnit(width);             // 4 byte
        w.writeHwpUnit(height);            // 4 byte
        w.writeInt32(zOrder);              // 4 byte
        for (int i = 0; i < 4; i++) {
            w.writeHwpUnit16(margins[i]);  // 총 8 byte
        }
        w.writeUInt32(instanceId);         // 4 byte
        w.writeInt32(pageBreakPrev);       // 4 byte
        // description 문자열
        String desc = (description != null) ? description : "";
        w.writeUInt16(desc.length());      // 2 byte
        w.writeUtf16Le(desc);              // 2*descLen byte
        return RecordWriter.buildRecord(HwpTagId.CTRL_HEADER, level, w.toByteArray());
    }

    /** double를 8 byte little-endian (IEEE 754)로 기록. */
    private static void writeDouble(HwpBinaryWriter w, double v) {
        long bits = Double.doubleToLongBits(v);
        w.writeUInt32(bits & 0xFFFFFFFFL);
        w.writeUInt32((bits >>> 32) & 0xFFFFFFFFL);
    }
}
