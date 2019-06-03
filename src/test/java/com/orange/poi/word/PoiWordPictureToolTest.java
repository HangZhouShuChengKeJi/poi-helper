package com.orange.poi.word;

import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;

import static org.junit.Assert.*;

/**
 * @author 小天
 * @date 2019/6/3 22:17
 */
public class PoiWordPictureToolTest {

    @Test
    public void createPicture() {
    }

    @Test
    public void addPicture() throws IOException, URISyntaxException {
        File img1 = new File(getClass().getResource("/img/linux_1.jpg").toURI());
        File img2 = new File(getClass().getResource("/img/linux_2.jpg").toURI());
        File img3 = new File(getClass().getResource("/img/linux_3.png").toURI());

        XWPFDocument doc = new XWPFDocument();
        PoiWordTool.initDocForA4(doc);

        // 图片自动缩放
        PoiWordPictureTool.addPicture(doc.createParagraph(), img1.getAbsolutePath(), true);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        PoiWordPictureTool.addPicture(doc.createParagraph(), img2.getAbsolutePath(), true);

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        PoiWordPictureTool.addPicture(doc.createParagraph(), img3.getAbsolutePath(), false);
        PoiWordParagraphTool.addBlankLine(doc);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    @Test
    public void getPictureType() {
    }

    @Test
    public void setPicturePosition() {
    }
}