package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.paper.PaperSize;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author 小天
 * @date 2019/6/2 23:30
 */
public class PoiWordToolTest {

    private String defaultFontFamily = "宋体";
    private int    defaultFontSize   = 14;
    private String defaultColor      = "000000";

    @Before
    public void setUp() throws Exception {
        File outputDir = new File("temp");
        System.setProperty("java.io.tmpdir", outputDir.getAbsolutePath());
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void initDocForA4() {

        System.out.println(PaperSize.A4.width);
        System.out.println(PaperSize.A4.height);

        System.out.println(PaperSize.A4.width_cm);
        System.out.println(PaperSize.A4.height_cm);

        System.out.println(PaperSize.A4.width_point);
        System.out.println(PaperSize.A4.height_point);

        System.out.println(PaperSize.B5.width_point);
        System.out.println(PaperSize.B5.height_point);
    }


    @Test
    public void createDoc() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();

        PoiWordTool.setDefaultStyle(doc, "Arial", "思源黑体 CN Light", 14, "FF0000");

        PoiWordParagraphTool.createParagraph(doc, "中文 ABC ______");

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void createDocA3H() throws IOException {
        // A4 横排版式，两栏测试

        XWPFDocument doc = PoiWordTool.createDoc();

        PoiWordTool.setDefaultStyle(doc, "Arial", "思源黑体 CN Light", 14, "FF0000");

        CTSectPr ctSectPr = PoiWordSectionTool.getSectPr(doc);
        PoiWordSectionTool.setPageSize(ctSectPr, PaperSize.A3_H, 20, 20, 15, 15);
        PoiWordSectionTool.setCols(ctSectPr, 2, 20, false);

        for (int i = 1; i <= 100; i++) {
            PoiWordParagraphTool.createParagraph(doc, "第 " + i + " 行：中文 ABC ______");
        }

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void addBlankLine() {
    }

    @Test
    public void createParagraph() {
    }

    @Test
    public void addParagraph() {
    }

    @Test
    public void setLineHeight() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);
        XWPFParagraph paragraph;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 1.5f);
        PoiWordParagraphTool.addTxt(paragraph, "1.5 倍 行间距测试", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 2.0f);
        PoiWordParagraphTool.addTxt(paragraph, "2.0 倍 行间距测试", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 3.0f);
        PoiWordParagraphTool.addTxt(paragraph, "3.0 倍 行间距测试", defaultFontFamily, defaultFontSize, defaultColor);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void setParagraphSpaceOfPound() {
    }

    @Test
    public void setParagraphSpaceOfLine() {
    }

    @Test
    public void addBreak() {
    }

    @Test
    public void setParagraphStyle() {
    }

    @Test
    public void getXWPFRun() {
    }

    @Test
    public void getRunProperties() {
    }

    @Test
    public void createPicture() {
    }

    @Test
    public void addPicture() {
    }

    @Test
    public void setPicturePosition() {
    }

    @Test
    public void createTableWithoutBorder() {
    }

    @Test
    public void createTable() {
    }

    @Test
    public void setTableRowHeightOfPixel() {
    }

    @Test
    public void setTableWidth() {
    }

    @Test
    public void setTableCellAlign() {
    }

    @Test
    public void setTableCellBgColor() {
    }

    @Test
    public void setPageMargin() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);

        // 设置页眉的位置
        PoiWordTool.setHeaderMargin(doc, PoiUnitTool.centimeterToDXA(0.6000f));
        // 设置页脚的位置
        PoiWordTool.setFooterMargin(doc, PoiUnitTool.centimeterToDXA(0.5000f));


        System.out.println(PoiWordTool.getContentWidthOfDxa(doc));


        XWPFParagraph paragraph;

        XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);
        paragraph = header.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "页眉", defaultFontFamily, defaultFontSize, defaultColor);

        XWPFFooter footer = doc.createFooter(HeaderFooterType.DEFAULT);
        paragraph = footer.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "页脚", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "文档内容", defaultFontFamily, defaultFontSize, defaultColor);




        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}