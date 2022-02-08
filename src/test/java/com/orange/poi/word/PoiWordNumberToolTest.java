package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import static org.junit.Assert.*;

/**
 * @author 小天
 * @date 2022/1/27 14:27
 */
public class PoiWordNumberToolTest {


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
    public void test() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();

        // 设置默认样式
        PoiWordTool.setDefaultStyle(doc, defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);

        XWPFParagraph paragraph;

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.addTxt(paragraph, "序号测试", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 0.5f, 0.5f);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);


        // 创建抽象编号
        XWPFAbstractNum xwpfAbstractNum = PoiWordNumberTool.createAbstractNumOfSingleLevel(doc);

        // 创建编号实例
        BigInteger numId = PoiWordNumberTool.createNumber(doc, xwpfAbstractNum);

        // 编号级别
        BigInteger numLevel = BigInteger.ZERO;

        // 编号级别样式
        PoiWordNumberTool.setLevel(xwpfAbstractNum, numLevel,
                1, STNumberFormat.DECIMAL, "%1.", STJc.CENTER,
                defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);

        double numLeft = PoiUnitTool.centimeterToPoint(0.300f);
        double textHanging = PoiUnitTool.centimeterToPoint(0.600f);

        // 设置编号位置 和 文本缩进
        PoiWordNumberTool.setIndByPoint(xwpfAbstractNum, numLevel, numLeft, textHanging);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        // 设置段落编号
        PoiWordParagraphTool.setNumber(paragraph, numId, numLevel);
        PoiWordParagraphTool.addTxt(paragraph, "第一段，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，", defaultFontFamily, defaultFontSize, defaultColor);

        // 一个编号有多个段时，后面的段需要通过设置左侧缩进来保持对齐。
        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.setLeftIndByPoint(paragraph, textHanging);
        PoiWordParagraphTool.addTxt(paragraph, "第二段. " +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.setLeftIndByPoint(paragraph, textHanging);
        PoiWordParagraphTool.addTxt(paragraph, "第三段. " +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进" +
                "段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进段落悬挂缩进", defaultFontFamily, defaultFontSize, defaultColor);


        paragraph = PoiWordParagraphTool.createParagraph(doc);
        // 设置段落编号
        PoiWordParagraphTool.setNumber(paragraph, numId, numLevel);
        PoiWordParagraphTool.addTxt(paragraph, "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.setLeftIndByPoint(paragraph, textHanging);
        PoiWordParagraphTool.addTxt(paragraph, "第二段. ", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        PoiWordParagraphTool.setLeftIndByPoint(paragraph, textHanging);
        PoiWordParagraphTool.addTxt(paragraph, "第三段. ", defaultFontFamily, defaultFontSize, defaultColor);


        paragraph = PoiWordParagraphTool.createParagraph(doc);
        // 设置段落编号
        PoiWordParagraphTool.setNumber(paragraph, numId, numLevel);
        PoiWordParagraphTool.addTxt(paragraph, "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，" +
                "编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，编号后面的文本，", defaultFontFamily, defaultFontSize, defaultColor);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        // 重新开始编号
        xwpfAbstractNum = PoiWordNumberTool.createAbstractNumOfSingleLevel(doc);
        numId = PoiWordNumberTool.createNumber(doc, xwpfAbstractNum);
        numLevel = BigInteger.ZERO;

        // 编号级别样式
        PoiWordNumberTool.setLevel(xwpfAbstractNum, numLevel,
                1, STNumberFormat.DECIMAL, "%1.", STJc.CENTER,
                defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        // 设置段落编号
        PoiWordParagraphTool.setNumber(paragraph, numId, numLevel);
        PoiWordParagraphTool.addTxt(paragraph, "重新开始编号", defaultFontFamily, defaultFontSize, defaultColor);

        paragraph = PoiWordParagraphTool.createParagraph(doc);
        // 设置段落编号
        PoiWordParagraphTool.setNumber(paragraph, numId, numLevel);
        PoiWordParagraphTool.addTxt(paragraph, "重新开始编号", defaultFontFamily, defaultFontSize, defaultColor);


        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}