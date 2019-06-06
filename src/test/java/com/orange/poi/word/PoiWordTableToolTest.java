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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

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
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void createTableWithoutBorder() throws URISyntaxException, IOException {
        File img3 = new File(getClass().getResource("/img/green_bg.png").toURI());

        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);

        XWPFParagraph paragraph;

        XWPFTable table = PoiWordTableTool.createTableWithoutBorder(doc, 1, 3);
        XWPFTableRow tableRowOne = table.getRow(0);
        // 设置表格高度
        PoiWordTableTool.setTableRowHeightOfPixel(tableRowOne, 40);

        XWPFTableCell tableCell = tableRowOne.getCell(0);
        PoiWordTableTool.setTableCellText(tableCell, "111111", STJc.LEFT, STVerticalJc.CENTER);

        tableCell = tableRowOne.getCell(1);
        PoiWordTableTool.setTableCellText(tableCell, "222222", STJc.LEFT, STVerticalJc.CENTER);

        tableCell = tableRowOne.getCell(2);
//        paragraph = tableCell.getParagraphArray(0);
        PoiWordTableTool.setTableCellText(tableCell, "33333", STJc.LEFT, STVerticalJc.CENTER);
//        PoiWordParagraphTool.addParagraph(paragraph, "33333", defaultFontFamily, defaultFontSize, defaultColor,
//                true, false,
//                null, TextAlignment.CENTER);
//        PoiWordParagraphTool.setLineHeightExact(paragraph, PoiUnitTool.pixelToPoint(40));

        tableCell = tableRowOne.getCell(0);
        tableCell.setColor("FF0000");
        paragraph = tableCell.getParagraphArray(0);
        PoiWordPictureTool.addPicture(paragraph, getClass().getResource("/img/bg.png").getFile());
        PoiWordPictureTool.setPicturePosition(paragraph,
                STRelFromH.MARGIN,0, STAlignH.LEFT,
                STRelFromV.PARAGRAPH, 0, null,
                true, false);

        // 第三列 背景图
        tableCell = tableRowOne.getCell(2);
        paragraph = tableCell.getParagraphArray(0);
        PoiWordPictureTool.addPicture(paragraph, img3);
        PoiWordPictureTool.setPicturePosition(paragraph,
                STRelFromH.MARGIN, (int) PoiUnitTool.pixelToPoint(452), null,
                STRelFromV.PARAGRAPH, (int) PoiUnitTool.pixelToPoint(7), null,
                true, false);

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
}