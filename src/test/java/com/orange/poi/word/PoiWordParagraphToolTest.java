package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRuby;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRubyContent;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRubyPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STRubyAlign;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URISyntaxException;

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
        File outputDir = new File("output");
        System.setProperty("java.io.tmpdir", outputDir.getAbsolutePath());
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void addBlankLine() {
    }

    @Test
    public void createParagraph() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.setSnapToGrid(paragraph, false);
        PoiWordParagraphTool.addTxt(paragraph, "段合并同类项，把结果按照x的降幂排列：2xy3-3xy2+（-5xy2）-（-8y2x）-（-3x2y）+4x3y．2xy3-3xy2+（-5xy2）-（-8y2x）-（-3x2y）+4x3y．2xy3-3xy2+（-5xy2）-（-8y2x）-（-3x2y）+4x3y．2xy3-3xy2+（-5xy2）-（-8y2x）-（-3x2y）+4x3y．2xy3-3xy2+（-5xy2）-（-8y2x）-（-3x2y）+4x3y．2xy3-3xy2+（-5xy2）-（-8y2x）-（-3x2y）+4x3y．", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 0.5f, 0.5f);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "新的段落", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 0.5f, 0.5f);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void addParagraph() {
    }

    @Test
    public void addSubscript() throws IOException, URISyntaxException, InvalidFormatException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph = PoiWordParagraphTool.createParagraph(doc);
        ;

        PoiWordParagraphTool.addTxt(paragraph, "y = log", defaultFontFamily, defaultFontSize, defaultColor);
//        PoiWordParagraphTool.addSubscript(paragraph, "2", defaultFontFamily, defaultFontSize, defaultColor);
//
//        paragraph = PoiWordParagraphTool.createParagraph(doc);
//        PoiWordParagraphTool.addTxt(paragraph, "3", defaultFontFamily, defaultFontSize, defaultColor);
//        PoiWordParagraphTool.addSuperscript(paragraph, "3", defaultFontFamily, defaultFontSize, defaultColor);


        File picFileIS = new File(getClass().getResource("/img/1.png").toURI());

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        ;
        XWPFRun paragraphRun = paragraph.createRun();
        paragraphRun.addPicture(new FileInputStream(picFileIS), XWPFDocument.PICTURE_TYPE_PNG, "123", 16, 26);
        if (StringUtils.isNotBlank(defaultFontFamily)) {
            paragraphRun.setFontFamily(defaultFontFamily);
        }
//        if (defaultFontSize != null) {
//            paragraphRun.setFontSize(defaultFontSize);
//        }
//        if (StringUtils.isNotBlank(color)) {
//            paragraphRun.setColor(color);
//        }
//        if (bold) {
//            paragraphRun.setBold(bold);
//        }
        paragraphRun.setSubscript(VerticalAlign.SUBSCRIPT);


        paragraph = PoiWordParagraphTool.createParagraph(doc);
        ;
        PoiWordParagraphTool.addTxt(paragraph, " x", defaultFontFamily, defaultFontSize, defaultColor);


        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void setLineHeightMultiple() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "1.0 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setLineHeightMultiple(paragraph, 0.5f);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
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
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "40磅 行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setLineHeightExact(paragraph, 40.00d);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
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
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "段前 0.5 倍行距，段后 0.5 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 0.5f, 0.5f);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "段前 1.0 倍行距，段后 1.0 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.0f, 1.0f);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "段前 1.5 倍行距，段后 1.5 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.5f, 1.5f);


        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void addBreak() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);

        PoiWordParagraphTool.addTxt(paragraph, "第一行");
        PoiWordParagraphTool.addBreak(paragraph);
        PoiWordParagraphTool.addTxt(paragraph, "第二行");
        PoiWordParagraphTool.addPageBreak(paragraph);
        PoiWordParagraphTool.addTxt(paragraph, "第二页，第一行");
        PoiWordParagraphTool.addPageBreak(doc);
        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "第三页，第一行");

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void addRuby() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addRuby(paragraph, "中", "zhong",
                "微软雅黑", 16, "000000",
                "微软雅黑", 8, "FF0000",
                5, "zh-CN");

        PoiWordParagraphTool.addRuby(paragraph, "国", "guo",
                "微软雅黑", 16, "000000",
                "微软雅黑", 8, "FF0000",
                5, "zh-CN");

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}