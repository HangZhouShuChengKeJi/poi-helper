package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;

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

    @Before
    public void setUp() throws Exception {
        File outputDir = new File("output");
        System.setProperty("java.io.tmpdir", outputDir.getAbsolutePath());
    }

    @Test
    public void createPicture() {
    }

    @Test
    public void addPicture() throws IOException, URISyntaxException {
        File img1 = new File(getClass().getResource("/img/1.png").toURI());

        XWPFDocument doc = PoiWordTool.createDocForA4();

        PoiWordParagraphTool.addTxt(doc.createParagraph(), "新的", defaultFontFamily, 25, defaultColor,
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

        XWPFDocument doc = PoiWordTool.createDocForA4();

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
    public void setPicturePosition() throws IOException, URISyntaxException {
        XWPFDocument doc = PoiWordTool.createDocForA4();

        File smallImgFile = new File(getClass().getResource("/img/2222.png").toURI());


        PoiWordParagraphTool.createParagraph(doc, "==== 左对齐 ====");

        XWPFParagraph xwpfParagraph = doc.createParagraph();

        PoiWordPictureTool.addPicture(xwpfParagraph, smallImgFile, 100, 100,
                STRelFromH.COLUMN, null, STAlignH.LEFT,
                STRelFromV.PARAGRAPH, (double) 0, null, null);

        PoiWordParagraphTool.addTxt(xwpfParagraph, "中旬，中央第五生态环境保护督察组通过“5+1”形式（5个下沉组+1个机动组），对郑州、安阳、新乡等十多个地市进行了下沉督察，对河南省各地区重点和问题线索进行更详细的梳理、追踪。4月18日，下沉工作刚刚开始，督察组就通过信访举报获悉，他们的车牌号已经被曝光：“如有企业人员看到‘京PXXXX’这个车牌号要注意，这是中央（第五）督察组的暗访车，一定及时在企业群里通知。”车牌号属于机动组。因为已经暴露，机动组下沉伊始就陷入了被动，仿佛身上被安装了GPS定位仪。在濮阳，3名督察组人员去濮阳检查一家羽绒加工企业，到附近排水口不到五分钟，几位县区相关部门的负责人跟上来“热情”地打招呼：“机动组领导来了？”“好家伙，您连机动组都知道了。”一位督察组人员说。兵分两路防跟踪车辆被跟踪，督察行程如何继续？机动组决定兵分两路，一拨人负责分散注意，另一拨人去现场锁定细节。第二天，机动组接到信访消息转道开封。在开封市精细化工产业集聚区，机动组迅速停车，去到信访件中提到的开封裕成化工有限公司。这家企业声称停产已久，但去年8月企业排水却被查出异常。在厂区，一位督察人员与企业负责人交谈，另一位督察人员去生产车间察看。值班室虽然空无一人，但桌上还放着一个沏了茶的水壶，茶水看起来并不像搁置了几个月的样子，最多只是前几天泡的。随后，督察人员攀爬到高处，发现废气治理设施仍在运转。督察组人员让企业立即出示生产记录，对方却说没有。十分钟后，企业工作人员有事离开，将原本手臂夹着的一个蓝色文件夹放在了一个不起眼的桌上，督察人员拿起来翻看，里面是车间几个月以来的生产设施用电、用蒸汽的付费明细。副镇长给企业直播督察进展这并不是中央生态环保督察组的行踪第一次被泄露。");

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);

        PoiWordParagraphTool.createParagraph(doc, "==== 右对齐 ====");

        xwpfParagraph = doc.createParagraph();

        PoiWordPictureTool.addPicture(xwpfParagraph, smallImgFile, 100, 100,
                STRelFromH.COLUMN, null, STAlignH.RIGHT,
                STRelFromV.PARAGRAPH, (double) 0, null, null);


        PoiWordParagraphTool.addTxt(xwpfParagraph, "中旬，中央第五生态环境保护督察组通过“5+1”形式（5个下沉组+1个机动组），对郑州、安阳、新乡等十多个地市进行了下沉督察，对河南省各地区重点和问题线索进行更详细的梳理、追踪。4月18日，下沉工作刚刚开始，督察组就通过信访举报获悉，他们的车牌号已经被曝光：“如有企业人员看到‘京PXXXX’这个车牌号要注意，这是中央（第五）督察组的暗访车，一定及时在企业群里通知。”车牌号属于机动组。因为已经暴露，机动组下沉伊始就陷入了被动，仿佛身上被安装了GPS定位仪。在濮阳，3名督察组人员去濮阳检查一家羽绒加工企业，到附近排水口不到五分钟，几位县区相关部门的负责人跟上来“热情”地打招呼：“机动组领导来了？”“好家伙，您连机动组都知道了。”一位督察组人员说。兵分两路防跟踪车辆被跟踪，督察行程如何继续？机动组决定兵分两路，一拨人负责分散注意，另一拨人去现场锁定细节。第二天，机动组接到信访消息转道开封。在开封市精细化工产业集聚区，机动组迅速停车，去到信访件中提到的开封裕成化工有限公司。这家企业声称停产已久，但去年8月企业排水却被查出异常。在厂区，一位督察人员与企业负责人交谈，另一位督察人员去生产车间察看。值班室虽然空无一人，但桌上还放着一个沏了茶的水壶，茶水看起来并不像搁置了几个月的样子，最多只是前几天泡的。随后，督察人员攀爬到高处，发现废气治理设施仍在运转。督察组人员让企业立即出示生产记录，对方却说没有。十分钟后，企业工作人员有事离开，将原本手臂夹着的一个蓝色文件夹放在了一个不起眼的桌上，督察人员拿起来翻看，里面是车间几个月以来的生产设施用电、用蒸汽的付费明细。副镇长给企业直播督察进展这并不是中央生态环保督察组的行踪第一次被泄露。");

        PoiWordParagraphTool.addBlankLine(doc);
        PoiWordParagraphTool.addBlankLine(doc);


        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}