package com.orange.poi.word;

import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.junit.Assert.*;

/**
 * @author 小天
 * @date 2019/6/5 10:59
 */
public class PoiWordParagraphToolTest {

    private String defaultFontFamily = "宋体";
    private int    defaultFontSize   = 14;
    private String defaultColor      = "000000";

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
    public void setLineHeightMultiple() {
    }

    @Test
    public void setLineHeightExact() {
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
        PoiWordParagraphTool.addParagraph(paragraph, "段前 0.5 倍行距，段后 0.5 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 0.5f, 0.5f);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addParagraph(paragraph, "段前 1.0 倍行距，段后 1.0 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.0f, 1.0f);

        paragraph = doc.createParagraph();
        PoiWordParagraphTool.addParagraph(paragraph, "段前 1.5 倍行距，段后 1.5 倍行距", defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.5f, 1.5f);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}