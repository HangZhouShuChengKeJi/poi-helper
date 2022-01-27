package com.orange.poi.word;

import com.microsoft.schemas.vml.CTGroup;
import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;

/**
 * @author 小天
 * @date 2021/5/18 10:44
 */
public class PoiTextBoxToolTest {
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
    public void createTextBox() throws IOException, URISyntaxException, InvalidFormatException {
        XWPFDocument xwpfDocument = PoiWordTool.createDocForA4();

        long contentWidth_dxa = PoiWordTool.getContentWidthOfDxa(xwpfDocument);

        File bgImg = new File(getClass().getResource("/img/title_bg.png").toURI());
        int pictureType = Document.PICTURE_TYPE_PNG;

        int width_px = 410;
        int height_px = 90;

        double width = PoiUnitTool.pixelToPoint(width_px);
        double height = PoiUnitTool.pixelToPoint(height_px);

        PoiWordTool.setDefaultStyle(xwpfDocument, "Arial", "思源黑体 CN Light", 14, "FF0000");


        // 创建文本框
        CTGroup ctGroup = PoiTextBoxTool.createTextBox(width, height, "left");

        // 设置背景图
        PoiTextBoxTool.setBackgroundImg(xwpfDocument, ctGroup, bgImg, pictureType);

        PoiTextBoxTool.setText(ctGroup, "第一讲",
                defaultFontFamily, defaultFontFamily, defaultFontSize, "000000",
                false, false, STJc.CENTER);

        PoiTextBoxTool.setParagraphSpaceOfPound(ctGroup, 15, 30, 0);

        XWPFParagraph paragraph = PoiWordParagraphTool.createParagraph(xwpfDocument);
        PoiTextBoxTool.addGroup(paragraph, ctGroup);

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        xwpfDocument.write(out);
        out.close();
    }

    @Test
    public void setBackgroundImg() {
    }

    @Test
    public void setText() {
    }
}