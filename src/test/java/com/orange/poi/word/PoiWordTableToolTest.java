package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STXAlign;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STYAlign;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;

/**
 * @author 小天
 * @date 2019/6/6 10:57
 */
public class PoiWordTableToolTest {

    private String defaultFontFamily = "宋体";
    private int    defaultFontSize   = 14;
    private String defaultColor      = "000000";

    @Before
    public void setUp() throws Exception {

        System.setProperty("java.io.tmpdir", System.getProperty("java.io.tmpdir") + "\\poiTest");
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void createTableWithoutBorder() throws URISyntaxException, IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);

        XWPFParagraph paragraph;

        XWPFTable table = PoiWordTableTool.createTable(doc, 1, 3, XWPFTable.XWPFBorderType.SINGLE, 2, "000000");
        XWPFTableRow tableRowOne = table.getRow(0);
        // 设置表格高度
        PoiWordTableTool.setTableRowHeightOfPixel(tableRowOne, 40);

        XWPFTableCell tableCell = tableRowOne.getCell(0);
        PoiWordTableTool.setTableCellText(tableCell, "对齐测试", STJc.LEFT, STVerticalJc.CENTER);

        tableCell = tableRowOne.getCell(1);
        PoiWordTableTool.setTableCellText(tableCell, "对齐测试测试", STJc.LEFT, STVerticalJc.CENTER);

        // 第三列
        tableCell = tableRowOne.getCell(2);
        paragraph = tableCell.getParagraphArray(0);

        PoiWordParagraphTool.addParagraph(paragraph, "对齐测试测", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordTableTool.setTableCellAlign(tableCell, STJc.LEFT, STVerticalJc.CENTER);

        PoiWordPictureTool.addPicture(paragraph, getClass().getResource("/img/star.png").getFile());
        PoiWordPictureTool.setPicturePosition(paragraph,
                STRelFromH.MARGIN, PoiUnitTool.pixelToPoint(100), null,
                STRelFromV.MARGIN, PoiUnitTool.pixelToPoint(5), null,
                true, true);

        // 添加空行
        PoiWordParagraphTool.addBlankLine(doc);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void createTable() {
    }

    @Test
    public void setTableRowHeightOfPixel() {
    }

    @Test
    public void setTableCellText() {
    }

    @Test
    public void setTableCellWidth() {
    }

    @Test
    public void setTableCellAlign() {
    }

    @Test
    public void setTableCellBgColor() {
    }

    @Test
    public void setTableCellBorderOfLeft() {

    }


    @Test
    public void setTablePosition() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);

        PoiWordParagraphTool.addParagraph(doc.createParagraph(), "右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格右边是表格", defaultFontFamily, defaultFontSize);


        XWPFTable table = PoiWordTableTool.createTableWithoutBorder(doc, 1, 3, true);
        XWPFTableRow tableRowOne = table.getRow(0);

        XWPFTableCell tableCell = tableRowOne.getCell(0);
        PoiWordTableTool.setTableCellText(tableCell, "111", STJc.LEFT, STVerticalJc.CENTER);

        PoiWordTableTool.setTableCellBorderOfBottom(tableCell, 1, "000000", STBorder.SINGLE);

        tableCell = tableRowOne.getCell(1);
        PoiWordTableTool.setTableCellText(tableCell, "222", STJc.LEFT, STVerticalJc.CENTER);

        // 第三列
        tableCell = tableRowOne.getCell(2);
        PoiWordTableTool.setTableCellText(tableCell, "333", STJc.LEFT, STVerticalJc.CENTER);

        PoiWordTableTool.setTablePosition(table, STHAnchor.TEXT, PoiUnitTool.pointToDXA(40),
                STVAnchor.TEXT, 0);
        PoiWordParagraphTool.addParagraph(doc.createParagraph(), "右边是表格", defaultFontFamily, defaultFontSize);
        PoiWordParagraphTool.addParagraph(doc.createParagraph(), "右边是表格", defaultFontFamily, defaultFontSize);
        // 添加空行
        PoiWordParagraphTool.addBlankLine(doc);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}