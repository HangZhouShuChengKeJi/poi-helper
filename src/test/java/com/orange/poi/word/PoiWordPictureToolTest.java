package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;

/**
 * @author 小天
 * @date 2019/6/3 22:17
 */
public class PoiWordPictureToolTest {

    private String defaultFontFamily = "宋体";
    private int    defaultFontSize   = 14;
    private String defaultColor      = "000000";

    @Test
    public void createPicture() {
    }

    @Test
    public void addPicture() throws IOException, URISyntaxException {
        File img1 = new File(getClass().getResource("/img/1.png").toURI());

        XWPFDocument doc =  PoiWordTool.createDocForA4();

        PoiWordParagraphTool.addTxt(doc.createParagraph(), "新的", defaultFontFamily , 25, defaultColor,
                true, false);

        // 设置背景图
        XWPFPicture picture1 = PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), img1, false);
        XWPFPicture picture2 = PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), img1, 500, 100, false);
        XWPFPicture picture3 = PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), img1, 300, -1, false);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }


    @Test
    public void addPictureWithResize() throws IOException, URISyntaxException {

        XWPFDocument doc =  PoiWordTool.createDocForA4();

//        // 添加图片
//        PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), new File(getClass().getResource("/img/1.jpg").toURI()), true);
//        PoiWordParagraphTool.addBlankLine(doc);
//        PoiWordParagraphTool.addBlankLine(doc);
//
//        // 添加图片
//        PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), new File(getClass().getResource("/img/1.jpg").toURI()), 1000, 1000, true);
//        PoiWordParagraphTool.addBlankLine(doc);
//        PoiWordParagraphTool.addBlankLine(doc);
//
//        PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), new File(getClass().getResource("/img/1.jpg").toURI()), 500, 500, true);
//        PoiWordParagraphTool.addBlankLine(doc);
//        PoiWordParagraphTool.addBlankLine(doc);

        PoiWordPictureTool.addPictureWithResize(doc.createParagraph(), new File(getClass().getResource("/img/3_1.jpg").toURI()), true);
        PoiWordParagraphTool.addBlankLine(doc);
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