package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.FileTypeEnum;
import com.orange.poi.util.FileTypeTool;
import com.orange.poi.util.FileUtil;
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
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTWrapSquare;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STAlignV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STWrapText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

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
        return addPicture(paragraph, new File(imgFile));
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
        BufferedImage image = ImageTool.readImage(imgFile);
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile.getAbsolutePath());
        }
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();
        return addPicture(paragraph, imgFile, actualWidth, actualHeight);
    }

    /**
     * 添加图片，当图片实际宽高超出内容区域的宽高时，对图片进行等比缩放。当图片溢出的时候，不通过重绘图片来缩小图片尺寸。缩放时，锁定原始长宽比例
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPictureWithResize(XWPFParagraph paragraph, File imgFile) throws IOException {
        return addPictureWithResize(paragraph, imgFile, false);
    }

    /**
     * 添加图片，当图片实际宽高超出内容区域的宽高时，对图片进行等比缩放。缩放时，锁定原始长宽比例
     *
     * @param paragraph        {@link XWPFParagraph}
     * @param imgFile          图片文件
     * @param redrawOnOverflow 当图片溢出的时候，是否通过重绘图片来缩小图片尺寸
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPictureWithResize(XWPFParagraph paragraph, File imgFile, boolean redrawOnOverflow) throws IOException {
        BufferedImage image = ImageTool.readImage(imgFile);
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile);
        }
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();
        // 使用原始长宽时，不需要再次指定 “锁定原始长宽比例”
        return addPictureWithResize(paragraph, imgFile, actualWidth, actualHeight, redrawOnOverflow, false);
    }

    /**
     * 添加图片，并指定宽高。当指定的宽高超出内容区域的宽高时，根据指定的宽高，对图片进行等比缩放。缩放时，锁定原始长宽比例
     *
     * @param paragraph        {@link XWPFParagraph}
     * @param imgFile          图片文件
     * @param width            宽度，单位：像素
     * @param height           高度，单位：像素
     * @param redrawOnOverflow 当图片溢出的时候，是否通过重绘图片来缩小图片尺寸
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPictureWithResize(XWPFParagraph paragraph, File imgFile, final int width, final int height, boolean redrawOnOverflow) throws IOException {
        return addPictureWithResize(paragraph, imgFile, width, height, redrawOnOverflow, true);
    }

    /**
     * 添加图片，并指定宽高。当指定的宽高超出内容区域的宽高时，根据指定的宽高，对图片进行等比缩放
     *
     * @param paragraph         {@link XWPFParagraph}
     * @param imgFile           图片文件
     * @param width             宽度，单位：像素
     * @param height            高度，单位：像素
     * @param redrawOnOverflow  当图片溢出的时候，是否通过重绘图片来缩小图片尺寸
     * @param lockOriginalScale 是否锁定原始长宽比例
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPictureWithResize(XWPFParagraph paragraph, File imgFile, final int width, final int height, boolean redrawOnOverflow, boolean lockOriginalScale) throws IOException {
        XWPFDocument doc = paragraph.getDocument();
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageSz pageSize = sectPr.getPgSz();
        if (pageSize == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) {
            throw new IllegalStateException("未设置页边距");
        }

        int contentWidth = PoiUnitTool.dxaToPixel((int) PoiWordTool.getContentWidthOfDxa(pageSize, pageMar));
        int contentHeight = PoiUnitTool.dxaToPixel((int) PoiWordTool.getContentHeightOfDxa(pageSize, pageMar));
        // 不能完全以contentWidth和contentHeight作为最大的宽高，需要留一定的空间 不然会出现大长图前一页空白 后一页空白的状况
        int delta = 0;
        if (width < height && height != 0) {
            float rate = width / (float) height;
            if (rate < 0.5) {
                delta = 200;
            } else if (rate < 0.65) {
                delta = 130;
            } else if (rate < 0.7) {
                delta = 110;
            } else {
                delta = 50;
            }
        }
        contentWidth -= delta;
        contentHeight -= delta;
        return addPictureWithResize(paragraph, imgFile, width, height, contentWidth, contentHeight, redrawOnOverflow, lockOriginalScale);
    }

    /**
     * 添加图片，并指定宽高。当指定的宽高超出指定的最大宽高时时，根据指定的宽高，对图片进行等比缩放
     *
     * @param paragraph         {@link XWPFParagraph}
     * @param imgFile           图片文件
     * @param width             宽度，单位：像素
     * @param height            高度，单位：像素
     * @param maxWidth          最大宽度，单位：像素
     * @param maxHeight         最大高度，单位：像素
     * @param redrawOnOverflow  当图片溢出的时候，是否通过重绘图片来缩小图片尺寸
     * @param lockOriginalScale 是否锁定原始长宽比例
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPictureWithResize(XWPFParagraph paragraph, File imgFile, final int width, final int height, final int maxWidth, final int maxHeight, boolean redrawOnOverflow, boolean lockOriginalScale) throws IOException {
        BufferedImage image = ImageTool.readImage(imgFile);
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile);
        }
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();

        int customizeWidth = width;
        int customizeHeight = height;
        // 宽高比例
        Double originalScale = (double) (1.0f * actualWidth / actualHeight);

        if (customizeWidth > 0 && customizeHeight <= 0) {
            // 当仅指定宽度时，根据原图宽高比例计算高度
            customizeHeight = (int) (customizeWidth / originalScale);
        }

        if (!lockOriginalScale) {
            originalScale = null;
        }

        if (customizeWidth > maxWidth || customizeHeight > maxHeight) {

            // 通过指定宽高，强制缩放图片
            final double scaleW = (double) customizeWidth / maxWidth;
            final double scaleH = (double) customizeHeight / maxHeight;
            if (scaleW > scaleH) {
                // 按照宽度缩放
                if (originalScale == null) {
                    return addPicture(paragraph, imgFile.getAbsolutePath(), maxWidth, (int) (customizeHeight / scaleW));
                } else {
                    return addPicture(paragraph, imgFile.getAbsolutePath(), maxWidth, (int) (maxWidth / originalScale));
                }
            }
            if (originalScale == null) {
                return addPicture(paragraph, imgFile.getAbsolutePath(), (int) (customizeWidth / scaleH), maxHeight);
            } else {
                return addPicture(paragraph, imgFile.getAbsolutePath(), (int) (maxHeight * originalScale), maxHeight);
            }
        }
        if (originalScale == null) {
            return addPicture(paragraph, imgFile.getAbsolutePath(), customizeWidth, customizeHeight);
        } else {
            return addPicture(paragraph, imgFile.getAbsolutePath(), customizeWidth, (int) (customizeWidth / originalScale));
        }
    }

    /**
     * 获取图片类型
     *
     * @param imgFile 图片文件
     * @param imgPath 图片路径
     *
     * @return 图片类型
     *
     * @see Document
     */
    public static Integer getPictureType(File imgFile, String imgPath) throws IOException {
        FileTypeEnum fileTypeEnum = FileTypeTool.getInstance().detect(imgFile);
        if (fileTypeEnum != null) {
            switch (fileTypeEnum) {
                case JPEG:
                    return XWPFDocument.PICTURE_TYPE_JPEG;
                case PNG:
                    return XWPFDocument.PICTURE_TYPE_PNG;
                case GIF:
                    return XWPFDocument.PICTURE_TYPE_GIF;
                case BMP:
                case BMP_16:
                case BMP_24:
                case BMP_256:
                    return XWPFDocument.PICTURE_TYPE_BMP;
                default:
            }
        }
        throw new IllegalArgumentException("不支持的文件格式: " + imgPath +
                ". 仅支持以下格式的图片： jpeg,png,gif,bmp");
    }

    /**
     * 添加图片（只负责基本的绘制操作，不做其他任何处理）
     *
     * @param paragraph {@link XWPFParagraph}
     * @param filePath  图片文件绝对地址
     * @param width     图片宽度（单位： 像素）
     * @param height    图片高度（单位： 像素）
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, String filePath, int width, int height) throws IOException {
        return addPicture(paragraph, new File(filePath), width, height);
    }

    /**
     * 添加图片（只负责基本的绘制操作，不做其他任何处理）
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件
     * @param width     图片宽度（单位： 像素）
     * @param height    图片高度（单位： 像素）
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, File imgFile, int width, int height) throws IOException {
        XWPFRun paragraphRun = paragraph.createRun();
        XWPFPicture picture = addPicture(paragraphRun, imgFile, width, height);
        return picture;
    }

    /**
     * 添加图片（只负责基本的绘制操作，不做其他任何处理）
     *
     * @param xwpfRun {@link XWPFRun}
     * @param imgFile 图片文件
     * @param width   图片宽度（单位： 像素）
     * @param height  图片高度（单位： 像素）
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFRun xwpfRun, File imgFile, int width, int height) throws IOException {
        XWPFPicture picture = null;

        Integer pictureType = getPictureType(imgFile, imgFile.getAbsolutePath());
        try {
            int width_emu = Units.pixelToEMU(width);
            int height_emu = Units.pixelToEMU(height);
            if (XWPFDocument.PICTURE_TYPE_JPEG == pictureType || XWPFDocument.PICTURE_TYPE_PNG == pictureType) {
                File tempFile = ImageTool.resetDensity(imgFile);
                if (tempFile == null) {
                    Integer density = ImageTool.getDensity(imgFile);
                    if (density != null && density == 150) {
                        // 像素密度为 150 时，需要将图片缩放
                        width_emu = (int) (width_emu * (96.0f / 150));
                        height_emu = (int) (height_emu * (96.0f / 150));
                    }
                    try (InputStream tempInputStream = FileUtil.readFile(imgFile)) {
                        picture = xwpfRun.addPicture(tempInputStream, pictureType, "", width_emu, height_emu);
                    }
                } else {
                    try (InputStream tempInputStream = FileUtil.readFile(tempFile)) {
                        picture = xwpfRun.addPicture(tempInputStream, pictureType, "", width_emu, height_emu);
                    }
                }
            } else {
                try (InputStream tempInputStream = FileUtil.readFile(imgFile)) {
                    picture = xwpfRun.addPicture(tempInputStream, pictureType, "", width_emu, height_emu);
                }
            }
            picture.getCTPicture().getSpPr().addNewNoFill();
            picture.getCTPicture().getSpPr().addNewLn().addNewNoFill();
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
    @Deprecated
    public static void setPicturePositionOfPage(XWPFParagraph paragraph, double leftOffset, double topOffset) {
        setPicturePosition(paragraph, STRelFromH.PAGE, leftOffset, null, STRelFromV.PAGE, topOffset, null,
                true, false, null);
    }

    /**
     * 设置图片相对页面边距的定位
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param leftOffset 水平偏移（单位： 磅）
     * @param topOffset  垂直偏移（单位： 磅）
     */
    @Deprecated
    public static void setPicturePositionOfPageMargin(XWPFParagraph paragraph, double leftOffset, double topOffset) {
        setPicturePosition(paragraph, STRelFromH.LEFT_MARGIN, leftOffset, null, STRelFromV.TOP_MARGIN, topOffset, null,
                true, false, null);
    }

    /**
     * 设置图片相对页面边距的定位
     *
     * @param paragraph {@link XWPFParagraph}
     * @param alignH    水平对齐方式
     * @param alignV    垂直对齐方式
     */
    @Deprecated
    public static void setPicturePositionOfPageMargin(XWPFParagraph paragraph, STAlignH.Enum alignH, STAlignV.Enum alignV) {
        setPicturePosition(paragraph, STRelFromH.MARGIN, null, alignH,
                STRelFromV.MARGIN, null, alignV, true, false, null);
    }

    /**
     * 设置图片相对段落的定位
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param leftOffset 水平偏移（单位： 磅）
     * @param topOffset  垂直偏移（单位： 磅）
     */
    @Deprecated
    public static void setPicturePositionOfParagraph(XWPFParagraph paragraph, double leftOffset, double topOffset, boolean layoutInCell) {
        setPicturePosition(paragraph, STRelFromH.COLUMN, leftOffset, null,
                STRelFromV.PARAGRAPH, topOffset, null, true, layoutInCell, null);
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
     * @param stWrapText            文字环绕方式
     */
    @Deprecated
    public static void setPicturePosition(XWPFParagraph paragraph,
                                          STRelFromH.Enum positionHRelativeFrom, Double leftOffset, STAlignH.Enum alignH,
                                          STRelFromV.Enum positionVRelativeFrom, Double topOffset, STAlignV.Enum alignV,
                                          boolean behindDoc, boolean layoutInCell,
                                          STWrapText.Enum stWrapText) {
        List<XWPFRun> runList = paragraph.getRuns();
        if (runList == null || runList.size() == 0) {
            return;
        }

        XWPFRun paragraphRun = runList.get(runList.size() - 1);
        setPicturePosition(paragraphRun,
                positionHRelativeFrom, leftOffset, alignH,
                positionVRelativeFrom, topOffset, alignV,
                behindDoc, layoutInCell,
                stWrapText);
    }

    /**
     * 设置图片位置
     *
     * @param paragraphRun          {@link XWPFRun}
     * @param positionHRelativeFrom 水平位置参考方式
     * @param leftOffset            水平偏移（单位： 磅）
     * @param alignH                水平位置对齐方式。仅当 positionHRelativeFrom 为 STRelFromH.MARGIN 时有用
     * @param positionVRelativeFrom 垂直位置参考方式
     * @param topOffset             垂直偏移（单位： 磅）
     * @param alignV                垂直位置对齐方式。仅当 positionVRelativeFrom 为 STRelFromV.MARGIN 时有用
     * @param behindDoc             是否置于文字底部
     * @param layoutInCell          是否在单元格内
     */
    public static void setPicturePosition(XWPFRun paragraphRun,
                                          STRelFromH.Enum positionHRelativeFrom, Double leftOffset, STAlignH.Enum alignH,
                                          STRelFromV.Enum positionVRelativeFrom, Double topOffset, STAlignV.Enum alignV,
                                          boolean behindDoc, boolean layoutInCell) {
        setPicturePosition(paragraphRun,
                positionHRelativeFrom, leftOffset, alignH,
                positionVRelativeFrom, topOffset, alignV,
                behindDoc, layoutInCell, null);
    }


    /**
     * 设置图片位置
     *
     * @param paragraphRun          {@link XWPFRun}
     * @param positionHRelativeFrom 水平位置参考方式
     * @param leftOffset            水平偏移（单位： 磅）
     * @param alignH                水平位置对齐方式。仅当 positionHRelativeFrom 为 STRelFromH.MARGIN 时有用
     * @param positionVRelativeFrom 垂直位置参考方式
     * @param topOffset             垂直偏移（单位： 磅）
     * @param alignV                垂直位置对齐方式。仅当 positionVRelativeFrom 为 STRelFromV.MARGIN 时有用
     * @param behindDoc             是否置于文字底部
     * @param layoutInCell          是否在单元格内
     * @param stWrapText            文字环绕方式
     */
    public static void setPicturePosition(XWPFRun paragraphRun,
                                          STRelFromH.Enum positionHRelativeFrom, Double leftOffset, STAlignH.Enum alignH,
                                          STRelFromV.Enum positionVRelativeFrom, Double topOffset, STAlignV.Enum alignV,
                                          boolean behindDoc, boolean layoutInCell,
                                          STWrapText.Enum stWrapText) {
        CTDrawing drawing = paragraphRun.getCTR().getDrawingArray(0);
        CTAnchor ctAnchor = drawing.addNewAnchor();

        ctAnchor.setSimplePos2(false);
        ctAnchor.setRelativeHeight(0);

        // 以下两个属性必须指定，否则使用 Microsoft Word 打开时，会提示文档已损坏
        ctAnchor.setLocked(false);
        ctAnchor.setLayoutInCell(layoutInCell);

        // 设置环绕方式
        if (stWrapText != null) {
            CTWrapSquare wrapSquare = ctAnchor.addNewWrapSquare();
            wrapSquare.setWrapText(stWrapText);
        }

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
