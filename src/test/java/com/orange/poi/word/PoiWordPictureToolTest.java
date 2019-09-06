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
        File img1 = new File(getClass().getResource("/img/linux_1.jpg").toURI());
        File img2 = new File(getClass().getResource("/img/linux_2.jpg").toURI());
        File img3 = new File(getClass().getResource("/img/linux_3.png").toURI());

        XWPFDocument doc =  PoiWordTool.createDocForA4();

        XWPFParagraph titleParagraph = doc.createParagraph();
        PoiWordParagraphTool.addTxt(titleParagraph, "背景图片测试", defaultFontFamily , 25, defaultColor,
                true, false);

        // 设置行高
        PoiWordParagraphTool.setLineHeightExact(titleParagraph, PoiUnitTool.pixelToPoint(200));

        // 设置背景图
        XWPFPicture picture = PoiWordPictureTool.addPicture(titleParagraph, img1);
        PoiWordPictureTool.setPicturePositionOfParagraph(titleParagraph, 0, 0, false);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }


    @Test
    public void addPictureWithResize() throws IOException, URISyntaxException {

        XWPFDocument doc =  PoiWordTool.createDocForA4();

        File img1 = new File(getClass().getResource("/img/1.jpg").toURI());
        XWPFParagraph paragraph = doc.createParagraph();

        // 添加图片
        PoiWordPictureTool.addPictureWithResize(paragraph, img1, false);

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