package com.orange.poi.word;

import com.orange.poi.util.TempFileUtil;
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
 * @date 2019/6/5 10:59
 */
public class PoiWordParagraphToolTest {

    private String defaultFontFamily = "思源宋体";
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
    public void addBlankLine() {
    }

    @Test
    public void createParagraph() {
    }

    @Test
    public void addParagraph() {
    }

    @Test
    public void addSubscript() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);
        XWPFParagraph paragraph;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "2", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.addSubscript(paragraph, "2", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "3", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.addSuperscript(paragraph, "3", defaultFontFamily, defaultFontSize, defaultColor);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void setLineHeightMultiple() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);
        XWPFParagraph paragraph;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "1.0 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 0.5f);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "2.0 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 2.0f);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void setLineHeightExact() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);
        XWPFParagraph paragraph;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "40磅 行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setLineHeightExact(paragraph, 40.00d);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "80磅 行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setLineHeightExact(paragraph, 80.00d);

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
    public void setParagraphSpaceOfLine() throws IOException {
        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);
        XWPFParagraph paragraph;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "段前 0.5 倍行距，段后 0.5 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 0.5f, 0.5f);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "段前 1.0 倍行距，段后 1.0 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.0f, 1.0f);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "段前 1.5 倍行距，段后 1.5 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.5f, 1.5f);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}