package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.ImageTool;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTEffectExtent;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTPosH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTPosV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import static com.orange.poi.word.PoiWordTool.A4_CONTENT_HEIGHT_DXA;
import static com.orange.poi.word.PoiWordTool.A4_CONTENT_WIDTH_DXA;

/**
 * apache poi word 图片工具类
 *
 * @author 小天
 * @date 2019/6/3 19:27
 */
public class PoiWordPictureTool {

    /**
     * 添加图片
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件绝对地址
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, String imgFile) throws IOException {
        return addPicture(paragraph, new File(imgFile), true);
    }

    /**
     * 添加图片
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, File imgFile) throws IOException {
        return addPicture(paragraph, imgFile, true);
    }

    /**
     * 添加图片
     *
     * @param paragraph        {@link XWPFParagraph}
     * @param imgFile          图片文件
     * @param redrawOnOverflow 当图片溢出的时候，是否通过重绘图片来缩小图片尺寸
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, File imgFile, boolean redrawOnOverflow) throws IOException {
        return addPicture(paragraph, imgFile, PoiUnitTool.dxaToPixel(A4_CONTENT_WIDTH_DXA), PoiUnitTool.dxaToPixel(A4_CONTENT_HEIGHT_DXA), redrawOnOverflow);
    }

    /**
     * 添加图片
     *
     * @param paragraph        {@link XWPFParagraph}
     * @param imgFile          图片文件
     * @param width            图片宽度（单位： 像素）
     * @param height           图片高度（单位： 像素）
     * @param redrawOnOverflow 当图片溢出的时候，是否通过重绘图片来缩小图片尺寸
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, File imgFile, int width, int height, boolean redrawOnOverflow) throws IOException {
        if (redrawOnOverflow) {
            ImageTool.ImageInfo imageInfo = ImageTool.resizeImage(imgFile, width, height);
            return addPicture(paragraph, imageInfo.getImgFile().getAbsolutePath(), imageInfo.getWidth(), imageInfo.getHeight());
        }
        BufferedImage image = ImageIO.read(imgFile);
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile);
        }
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();
        final double scaleW = (double) actualWidth / width;
        final double scaleH = (double) actualHeight / height;
        if (scaleW > scaleH) {
            return addPicture(paragraph, imgFile.getAbsolutePath(), width, (int) (actualHeight / scaleW));
        }
        return addPicture(paragraph, imgFile.getAbsolutePath(), (int) (width / scaleH), height);
    }

    /**
     * 获取图片类型
     *
     * @param imgFile 图片文件名称
     *
     * @return 图片类型
     *
     * @see Document
     */
    public static Integer getPictureType(String imgFile) {
        if (imgFile.endsWith(".jpg") || imgFile.endsWith(".jpeg")) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        } else if (imgFile.endsWith(".png")) {
            return XWPFDocument.PICTURE_TYPE_PNG;
        } else if (imgFile.endsWith(".emf")) {
            return XWPFDocument.PICTURE_TYPE_EMF;
        } else if (imgFile.endsWith(".wmf")) {
            return XWPFDocument.PICTURE_TYPE_WMF;
        } else if (imgFile.endsWith(".pict")) {
            return XWPFDocument.PICTURE_TYPE_PICT;
        } else if (imgFile.endsWith(".dib")) {
            return XWPFDocument.PICTURE_TYPE_DIB;
        } else if (imgFile.endsWith(".gif")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        } else if (imgFile.endsWith(".tiff")) {
            return XWPFDocument.PICTURE_TYPE_TIFF;
        } else if (imgFile.endsWith(".eps")) {
            return XWPFDocument.PICTURE_TYPE_EPS;
        } else if (imgFile.endsWith(".bmp")) {
            return XWPFDocument.PICTURE_TYPE_BMP;
        } else if (imgFile.endsWith(".wpg")) {
            return XWPFDocument.PICTURE_TYPE_WPG;
        } else {
            throw new IllegalArgumentException("Unsupported picture: " + imgFile +
                    ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
        }
    }

    /**
     * 添加图片（只负责基本的绘制操作，不做其他任何处理）
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件绝对地址
     * @param width     图片宽度（单位： 像素）
     * @param height    图片高度（单位： 像素）
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, String imgFile, int width, int height) throws IOException {
        XWPFRun paragraphRun = paragraph.createRun();
        XWPFPicture picture = null;

        try (FileInputStream is = new FileInputStream(imgFile)) {
            picture = paragraphRun.addPicture(is, getPictureType(imgFile), imgFile, Units.pixelToEMU(width), Units.pixelToEMU(height));
        } catch (InvalidFormatException ignore) {
        }
        return picture;
    }


    /**
     * 添加图片（只负责基本的绘制操作，不做其他任何处理）
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件绝对地址
     * @param width     图片宽度（单位： 像素）
     * @param height    图片高度（单位： 像素）
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, File imgFile, int width, int height) throws IOException {
        XWPFRun paragraphRun = paragraph.createRun();
        XWPFPicture picture = null;

        try (FileInputStream is = new FileInputStream(imgFile)) {
            picture = paragraphRun.addPicture(is, getPictureType(imgFile.getAbsolutePath()), imgFile.getAbsolutePath(), Units.pixelToEMU(width), Units.pixelToEMU(height));
        } catch (InvalidFormatException ignore) {
        }
        return picture;
    }

    /**
     * 设置图片相对页面的定位
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param leftOffset 水平偏移（单位： 磅）
     * @param topOffset  垂直偏移（单位： 磅）
     */
    public static void setPicturePositionOfPage(XWPFParagraph paragraph, int leftOffset, int topOffset) {
        setPicturePosition(paragraph, STRelFromH.PAGE, leftOffset, null, STRelFromV.PAGE, topOffset, null,
                true, false);
    }

    /**
     * 设置图片相对页面边距的定位
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param leftOffset 水平偏移（单位： 磅）
     * @param topOffset  垂直偏移（单位： 磅）
     */
    public static void setPicturePositionOfPageMargin(XWPFParagraph paragraph, int leftOffset, int topOffset) {
        setPicturePosition(paragraph, STRelFromH.LEFT_MARGIN, leftOffset, null, STRelFromV.TOP_MARGIN, topOffset, null,
                true, false);
    }

    /**
     * 设置图片相对页面边距的定位
     *
     * @param paragraph {@link XWPFParagraph}
     * @param alignH    水平对齐方式
     * @param alignV    垂直对齐方式
     */
    public static void setPicturePositionOfPageMargin(XWPFParagraph paragraph, STAlignH.Enum alignH, STAlignV.Enum alignV) {
        setPicturePosition(paragraph, STRelFromH.MARGIN, 0, alignH,
                STRelFromV.MARGIN, 0, alignV, true, false);
    }

    /**
     * 设置图片相对段落的定位
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param leftOffset 水平偏移（单位： 磅）
     * @param topOffset  垂直偏移（单位： 磅）
     */
    public static void setPicturePositionOfParagraph(XWPFParagraph paragraph, int leftOffset, int topOffset, boolean layoutInCell) {
        setPicturePosition(paragraph, STRelFromH.COLUMN, leftOffset, null,
                STRelFromV.PARAGRAPH, topOffset, null, true, layoutInCell);
    }

    /**
     * 设置图片位置
     *
     * @param paragraph             {@link XWPFParagraph}
     * @param positionHRelativeFrom 水平位置参考方式
     * @param leftOffset            水平偏移（单位： 磅）
     * @param alignH                水平位置对齐方式。仅当 positionHRelativeFrom 为 STRelFromH.MARGIN 时有用
     * @param positionVRelativeFrom 垂直位置参考方式
     * @param topOffset             垂直偏移（单位： 磅）
     * @param alignV                垂直位置对齐方式。仅当 positionVRelativeFrom 为 STRelFromV.MARGIN 时有用
     * @param behindDoc             是否置于文字底部
     * @param layoutInCell          是否在单元格内
     */
    public static void setPicturePosition(XWPFParagraph paragraph,
                                          STRelFromH.Enum positionHRelativeFrom, Integer leftOffset, STAlignH.Enum alignH,
                                          STRelFromV.Enum positionVRelativeFrom, Integer topOffset, STAlignV.Enum alignV,
                                          boolean behindDoc, boolean layoutInCell) {
        List<XWPFRun> runList = paragraph.getRuns();
        if (runList == null || runList.size() == 0) {
            return;
        }

        XWPFRun paragraphRun = runList.get(runList.size() - 1);
        CTDrawing drawing = paragraphRun.getCTR().getDrawingArray(0);
        CTAnchor ctAnchor = drawing.addNewAnchor();

        ctAnchor.setSimplePos2(false);
        ctAnchor.setRelativeHeight(0);

        // 以下两个属性必须指定，否则使用 Microsoft Word 打开时，会提示文档已损坏
        ctAnchor.setLocked(false);
        ctAnchor.setLayoutInCell(layoutInCell);

        // 水平位置
        CTPosH posH;
        if ((posH = ctAnchor.getPositionH()) == null) {
            posH = ctAnchor.addNewPositionH();
        }
        if (positionHRelativeFrom != null) {
            posH.setRelativeFrom(positionHRelativeFrom);
            if (leftOffset != null) {
                posH.setPosOffset(Units.toEMU(leftOffset));
            } else {
                posH.setAlign(alignH);
            }
        }

        // 垂直位置
        CTPosV posV;
        if ((posV = ctAnchor.getPositionV()) == null) {
            posV = ctAnchor.addNewPositionV();
        }
        if (positionVRelativeFrom != null) {
            posV.setRelativeFrom(positionVRelativeFrom);
            if (topOffset != null) {
                posV.setPosOffset(Units.toEMU(topOffset));
            } else {
                posV.setAlign(alignV);
            }
        }

        // 复制原有的属性
        CTInline ctInline = drawing.getInlineArray(0);

        ctAnchor.setDistT(ctInline.getDistT());
        ctAnchor.setDistR(ctInline.getDistR());
        ctAnchor.setDistB(ctInline.getDistB());
        ctAnchor.setDistL(ctInline.getDistL());

        // 置于文字底部
        ctAnchor.setBehindDoc(behindDoc);
        if (behindDoc) {
            ctAnchor.addNewWrapNone();
        }

        // 允许图片叠加
        ctAnchor.setAllowOverlap(true);

        ctAnchor.setExtent(ctInline.getExtent());
        ctAnchor.setDocPr(ctInline.getDocPr());

        CTPoint2D simplePos = ctAnchor.addNewSimplePos();
        simplePos.setX(0);
        simplePos.setY(0);

        CTEffectExtent effectExtent = ctAnchor.addNewEffectExtent();
        effectExtent.setT(0);
        effectExtent.setR(0);
        effectExtent.setB(0);
        effectExtent.setL(0);

        ctAnchor.addNewCNvGraphicFramePr();

        // 将旧的图片数据拷贝过来
        ctAnchor.setGraphic(ctInline.getGraphic());


        // 移除旧的图片
        drawing.removeInline(0);
    }
}
