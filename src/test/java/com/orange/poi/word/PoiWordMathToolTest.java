package com.orange.poi.word;

import com.orange.poi.util.OfficeMathMLUtil;
import com.orange.poi.util.TempFileUtil;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;
import org.dom4j.DocumentException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import javax.xml.transform.TransformerException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

/**
 * @author 小天
 * @date 2019/7/5 23:23
 */
public class PoiWordMathToolTest {

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
    public void addMathML() throws IOException, XmlException, DocumentException, TransformerException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        String mml;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "公式——分数：", defaultFontFamily, defaultFontSize, defaultColor);
        mml = IOUtils.resourceToString("/mml/mml-frac.xml", StandardCharsets.UTF_8);
        PoiWordMathTool.addMathML(paragraph, mml);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "公式——开根号：", defaultFontFamily, defaultFontSize, defaultColor);
        mml = IOUtils.resourceToString("/mml/mml-sqrt.xml", StandardCharsets.UTF_8);
        PoiWordMathTool.addMathML(paragraph, mml);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "公式——方程组(mfenced)：", defaultFontFamily, defaultFontSize, defaultColor);
        mml = IOUtils.resourceToString("/mml/mml-mfenced.xml", StandardCharsets.UTF_8);
        PoiWordMathTool.addMathML(paragraph, mml);


        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "公式——方程组(mathjax-env-cases)：", defaultFontFamily, defaultFontSize, defaultColor);
        mml = IOUtils.resourceToString("/mml/mml-env-cases.xml", StandardCharsets.UTF_8);
        PoiWordMathTool.addMathML(paragraph, mml);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void addFraction() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        XWPFParagraph paragraph;

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "x=", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordMathTool.addFraction(paragraph, "a", "b", defaultFontFamily, "ff0000");


        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(paragraph, "计算： x=", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordMathTool.addFraction(paragraph, "999", "1000", defaultFontFamily, "00dd00");

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void convertMmlToOmml() throws TransformerException, DocumentException {

        String mathml = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">\n  <mfrac>\n    <mn>1</mn>\n    <mn>2</mn>\n  </mfrac>\n</math>";
        String oMathML = OfficeMathMLUtil.getInstance().convertMmlToOmml(mathml);
        System.out.println(oMathML);
    }

    @Test
    public void convertOmmlToMml() {

    }
}