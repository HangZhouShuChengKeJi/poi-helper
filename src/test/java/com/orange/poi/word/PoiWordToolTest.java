package com.orange.poi.word;

import com.orange.poi.util.TempFileUtil;
import com.orange.poi.word.PoiWordParagraphTool;
import com.orange.poi.word.PoiWordTool;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

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
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void initDocForA4() {
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
        PoiWordParagraphTool.addParagraph(paragraph, "1.5 倍 行间距测试", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 2.0f);
        PoiWordParagraphTool.addParagraph(paragraph, "2.0 倍 行间距测试", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 3.0f);
        PoiWordParagraphTool.addParagraph(paragraph, "3.0 倍 行间距测试", defaultFontFamily, defaultFontSize, defaultColor);

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
}