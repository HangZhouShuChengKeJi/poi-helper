package com.orange.poi;

import com.orange.poi.util.TempFileUtil;
import com.orange.poi.util.UrlUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTEffectExtent;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTPosH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTPosV;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.List;

/**
 * apache poi word 工具类
 *
 * @author 小天
 * @date 2019/5/27 22:23
 * @see <a href="http://poi.apache.org/">Apache POI</a>
 */
public class PoiWordTool {
    /**
     * 默认字体大小，单位：磅
     */
    public static final  int DEFAULT_FONT_SIZE = 12;
    /**
     * 常量，在文档中定义长度和高度的单位
     */
    private static final int PER_LINE          = 100;
    /**
     * 基准行高
     */
    public static final  int LINE_HEIGHT_DXA   = 240;

    /**
     * A4 纸宽度，单位：DXA
     */
    public static final int  A4_WIDTH_DXA          = 11906;
    /**
     * A4 纸高度，单位：DXA
     */
    public static final int  A4_HEIGHT_DXA         = 16838;
    /**
     * 页面上边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_TOP_DXA     = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 页面右边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_RIGHT_DXA   = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 页面下边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_BOTTOM_DXA  = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 页面左边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_LEFT_DXA    = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 内容宽度（页面宽度减去左右页边距），单位：DXA
     */
    public static final int  A4_CONTENT_WIDTH_DXA  = A4_WIDTH_DXA - A4_MARGIN_LEFT_DXA - A4_MARGIN_RIGHT_DXA;
    /**
     * 内容高度（页面高度减去上下页边距），单位：DXA
     */
    public static final int  A4_CONTENT_HEIGHT_DXA = A4_HEIGHT_DXA - A4_MARGIN_TOP_DXA - A4_MARGIN_BOTTOM_DXA;
    /**
     * 临时文件目录
     */
    public static final File TEMP_FILE_DIR         = new File(System.getProperty("java.io.tmpdir"));

    /**
     * 将 word 初始化为 A4 值（包括页面大小、页边距等）
     */
    public static void initDocForA4(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageSz pageSize = sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(A4_WIDTH_DXA));
        pageSize.setH(BigInteger.valueOf(A4_HEIGHT_DXA));

        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setTop(BigInteger.valueOf(A4_MARGIN_TOP_DXA));
        pageMar.setRight(BigInteger.valueOf(A4_MARGIN_RIGHT_DXA));
        pageMar.setBottom(BigInteger.valueOf(A4_MARGIN_BOTTOM_DXA));
        pageMar.setLeft(BigInteger.valueOf(A4_MARGIN_LEFT_DXA));
    }


    /**
     * 创建空行
     *
     * @param document {@link XWPFDocument}
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph addBlankLine(XWPFDocument document) {
        return createParagraph(document, "", null, null, null);
    }


    /**
     * 创建段落
     *
     * @param document   {@link XWPFDocument}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, String plainTxt,
                                                String fontFamily, Integer fontSize, String color) {
        XWPFParagraph paragraph = document.createParagraph();
        addParagraph(paragraph, plainTxt, fontFamily, fontSize, color, false, false, ParagraphAlignment.LEFT, TextAlignment.CENTER);
        return paragraph;
    }

    /**
     * 创建段落
     *
     * @param document      {@link XWPFDocument}
     * @param plainTxt      文本内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param color         颜色（RGB 格式，例如："FFFFFF"）
     * @param bold          是否加粗
     * @param underline     是否增加下划线
     * @param alignment     水平对齐
     * @param verticalAlign 垂直对其
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, String plainTxt,
                                                String fontFamily, Integer fontSize, String color,
                                                boolean bold, boolean underline,
                                                ParagraphAlignment alignment, TextAlignment verticalAlign) {
        XWPFParagraph paragraph = document.createParagraph();
        addParagraph(paragraph, plainTxt, fontFamily, fontSize, color, bold, underline, alignment, verticalAlign);
        return paragraph;
    }

    /**
     * 添加段落内容
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void addParagraph(XWPFParagraph paragraph, String plainTxt,
                                    String fontFamily, Integer fontSize, String color) {
        addParagraph(paragraph, plainTxt, fontFamily, fontSize, color, false, false, ParagraphAlignment.LEFT, TextAlignment.CENTER);
    }


    /**
     * 添加段落内容
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     * @param bold       是否加粗
     * @param underline  是否增加下划线
     */
    public static void addParagraph(XWPFParagraph paragraph, String plainTxt,
                                    String fontFamily, Integer fontSize, String color,
                                    boolean bold, boolean underline) {
        addParagraph(paragraph, plainTxt, fontFamily, fontSize, color, bold, underline, ParagraphAlignment.LEFT, TextAlignment.CENTER);
    }

    /**
     * 添加段落内容
     *
     * @param paragraph     {@link XWPFParagraph}
     * @param plainTxt      文本内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param color         颜色（RGB 格式，例如："FFFFFF"）
     * @param bold          是否加粗
     * @param underline     是否增加下划线
     * @param alignment     水平对齐
     * @param verticalAlign 垂直对其
     */
    public static void addParagraph(XWPFParagraph paragraph, String plainTxt,
                                    String fontFamily, Integer fontSize, String color,
                                    boolean bold, boolean underline,
                                    ParagraphAlignment alignment, TextAlignment verticalAlign) {
        if (paragraph == null) {
            return;
        }
        XWPFRun paragraphRun = paragraph.createRun();
        paragraphRun.setText(plainTxt);
        if (StringUtils.isNotBlank(fontFamily)) {
            paragraphRun.setFontFamily(fontFamily);
        }
        if (fontSize != null) {
            paragraphRun.setFontSize(fontSize);
        }
        if (StringUtils.isNotBlank(color)) {
            paragraphRun.setColor(color);
        }
        paragraphRun.setBold(bold);
        if (underline) {
            paragraphRun.setUnderline(UnderlinePatterns.SINGLE);
        }
        if (alignment != null) {
            paragraph.setAlignment(alignment);
        }
        if (verticalAlign != null) {
            paragraph.setVerticalAlignment(verticalAlign);
        }
    }

    /**
     * 设置行高
     *
     * @param paragraph {@link XWPFParagraph}
     * @param multiple  多倍行距，例如： 1.5f 表示 1.5 倍行距
     */
    public static void setLineHeightMultiple(XWPFParagraph paragraph, double multiple) {
        if (multiple == 1.0f) {
            return;
        }
        CTPPr ppr = getParagraphProperties(paragraph);
        CTSpacing spacing;
        if ((spacing = ppr.getSpacing()) == null) {
            spacing = ppr.addNewSpacing();
        }
        spacing.setLine(BigInteger.valueOf((long) (multiple * LINE_HEIGHT_DXA)));
        spacing.setLineRule(STLineSpacingRule.AUTO);
    }

    /**
     * 设置行高
     *
     * @param paragraph {@link XWPFParagraph}
     * @param value     行高，单位：磅
     */
    public static void setLineHeightExact(XWPFParagraph paragraph, double value) {
        CTPPr ppr = getParagraphProperties(paragraph);
        CTSpacing spacing;
        if ((spacing = ppr.getSpacing()) == null) {
            spacing = ppr.addNewSpacing();
        }
        spacing.setLine(PoiUnitTool.pointToDXA(value));
        spacing.setLineRule(STLineSpacingRule.EXACT);
    }

    /**
     * 获取段落属性
     *
     * @param paragraph 段落 {@link XWPFParagraph}
     *
     * @return 段落属性 {@link CTPPr}
     */
    public static CTPPr getParagraphProperties(XWPFParagraph paragraph) {
        CTPPr ppr;
        if ((ppr = paragraph.getCTP().getPPr()) == null) {
            return paragraph.getCTP().addNewPPr();
        }
        return ppr;
    }

    /**
     * 设置段落间距（以磅为单位）
     *
     * @param paragraph {@link XWPFParagraph}
     * @param before    段落前间距（单位：磅）
     * @param after     段落后间距（单位：磅）
     */
    public static void setParagraphSpaceOfPound(XWPFParagraph paragraph, double before, double after) {
        CTPPr ppr = getParagraphProperties(paragraph);
        CTSpacing spacing;
        if ((spacing = ppr.getSpacing()) == null) {
            spacing = ppr.addNewSpacing();
        }
        spacing.setBeforeLines(PoiUnitTool.pointToDXA(before));
        spacing.setAfterLines(PoiUnitTool.pointToDXA(after));
    }

    /**
     * 设置段落间距（以行为单位）
     *
     * @param paragraph   {@link XWPFParagraph}
     * @param beforeLines 段落前间距（单位：行，例如：1.5f 表示 1.5 倍行距）
     * @param afterLines  段落后间距（单位：行，例如：1.5f 表示 1.5 倍行距）
     */
    public static void setParagraphSpaceOfLine(XWPFParagraph paragraph, double beforeLines, double afterLines) {
        CTPPr ppr = getParagraphProperties(paragraph);
        CTSpacing spacing;
        if ((spacing = ppr.getSpacing()) == null) {
            spacing = ppr.addNewSpacing();
        }
        spacing.setBefore(BigInteger.valueOf((long) (beforeLines * LINE_HEIGHT_DXA)));
        spacing.setAfterLines(BigInteger.valueOf((long) (afterLines * LINE_HEIGHT_DXA)));
    }

    /**
     * 添加回车符（不产生新的段落）
     *
     * @param paragraph {@link XWPFParagraph}
     */
    public static void addBreak(XWPFParagraph paragraph) {
        XWPFRun paragraphRun = getXWPFRun(paragraph);
        paragraphRun.addBreak();
    }

    /**
     * @param paragraph {@link XWPFParagraph}
     *
     * @return {@link XWPFRun}
     */
    private static XWPFRun getLastXWPFRun(XWPFParagraph paragraph) {
        XWPFRun paragraphRun = null;
        if (paragraph.getRuns() != null && paragraph.getRuns().size() > 0) {
            paragraphRun = paragraph.getRuns().get(paragraph.getRuns().size() - 1);
        }
        return paragraphRun;
    }

    /**
     * @param paragraph {@link XWPFParagraph}
     *
     * @return {@link XWPFRun}
     */
    public static XWPFRun getXWPFRun(XWPFParagraph paragraph) {
        XWPFRun paragraphRun = null;
        if (paragraph.getRuns() != null && paragraph.getRuns().size() > 0) {
            paragraphRun = paragraph.getRuns().get(0);
        } else {
            paragraphRun = paragraph.createRun();
        }
        return paragraphRun;
    }

    /**
     * 获取段落属性
     *
     * @param paragraph {@link XWPFParagraph}
     * @param create    true: 属性不存在时创建，否则不创建
     *
     * @return 段落属性，如果没有或者创建失败时返回 null
     */
    public static CTRPr getRunProperties(XWPFParagraph paragraph, boolean create) {
        XWPFRun paragraphRun = getXWPFRun(paragraph);
        if (paragraphRun == null) {
            return null;
        }
        CTR run = paragraphRun.getCTR();
        CTRPr pr = run.isSetRPr() ? run.getRPr() : null;
        if (create && pr == null) {
            pr = run.addNewRPr();
        }
        return pr;
    }

    /**
     * 创建图片
     *
     * @param document {@link XWPFDocument}
     * @param imgFile  图片文件绝对地址
     * @param width    图片宽度（单位： 像素）
     * @param height   图片高度（单位： 像素）
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture createPicture(XWPFDocument document, String imgFile, int width, int height) throws IOException {
        XWPFParagraph paragraph = document.createParagraph();
        return addPicture(paragraph, imgFile, width, height);
    }

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
        BufferedImage image = ImageIO.read(new File(imgFile));
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile);
        }
        int height = image.getHeight();
        int width = image.getWidth();
        return addPicture(paragraph, imgFile, width, height);
    }

    /**
     * 添加图片
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param imgFile    图片文件绝对地址
     * @param autoResize 自动调整图片大小
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, String imgFile, boolean autoResize) throws IOException {
        BufferedImage image = ImageIO.read(new File(imgFile));
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile);
        }
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();
        if (autoResize) {
            final int actualWidthDXA = PoiUnitTool.pixelToDXA(actualWidth);
            final int actualHeightDXA = PoiUnitTool.pixelToDXA(actualHeight);
            if (actualWidthDXA > A4_CONTENT_WIDTH_DXA || actualHeightDXA > A4_CONTENT_HEIGHT_DXA) {
                final double scaleW = actualWidth / A4_CONTENT_WIDTH_DXA;
                final double scaleH = actualHeight / A4_CONTENT_HEIGHT_DXA;
                final double scale = NumberUtils.max(scaleW, scaleH);

                final int newWidth = (int) (actualWidthDXA / scale);
                final int newHeight = (int) (actualHeightDXA / scale);

                final String imgEx = UrlUtil.getExNameFromUrl(imgFile);

                BufferedImage newImage = new BufferedImage(newWidth, newHeight, BufferedImage.TYPE_INT_RGB);
                newImage.getGraphics().drawImage(image, 0, 0, newWidth, newHeight, null);

                File newImgFile = TempFileUtil.createTempFile(imgEx);

                try (FileOutputStream fileOutputStream = new FileOutputStream(newImgFile)) {
                    ImageIO.write(newImage, imgEx, fileOutputStream);
                    return addPicture(paragraph, newImgFile.getAbsolutePath(), newWidth, newHeight);
                }
            }
        }
        return addPicture(paragraph, imgFile, actualWidth, actualHeight);
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
            picture = paragraphRun.addPicture(is, getPictureType(imgFile), imgFile, Units.toEMU(width), Units.toEMU(height));
        } catch (InvalidFormatException ignore) {
        }
        return picture;
    }

    /**
     * 添加图片
     *
     * @param paragraph {@link XWPFParagraph}
     * @param imgFile   图片文件绝对地址
     * @param width     图片宽度（单位： 像素）
     * @param height    图片高度（单位： 像素）
     * @param redraw    是否通过重绘的方式缩放图片
     *
     * @return {@link XWPFPicture}
     *
     * @throws IOException
     */
    public static XWPFPicture addPicture(XWPFParagraph paragraph, String imgFile, int width, int height, boolean redraw) throws IOException {
        XWPFRun paragraphRun = paragraph.createRun();
        XWPFPicture picture = null;
        int format = getPictureType(imgFile);

        final BufferedImage image = ImageIO.read(new File(imgFile));
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile);
        }

        final int expectedWidth0 = PoiUnitTool.pixelToDXA(width);
        final int expectedHeight0 = PoiUnitTool.pixelToDXA(height);
        final double scaleW0 = expectedWidth0 / A4_CONTENT_WIDTH_DXA;
        final double scaleH0 = expectedHeight0 / A4_CONTENT_HEIGHT_DXA;
        final double scale0 = NumberUtils.min(scaleW0, scaleH0);

        final int expectedWidth;
        final int expectedHeight;
        if (scale0 > 1.0f) {
            expectedWidth = (int) (expectedWidth0 / scale0);
            expectedHeight = (int) (expectedHeight0 / scale0);
        } else {
            expectedWidth = expectedWidth0;
            expectedHeight = expectedHeight0;
        }

        // 图片实际尺寸
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();

        if (redraw) {
            //构建图片流
            BufferedImage tag = new BufferedImage(expectedWidth, expectedHeight, BufferedImage.TYPE_INT_RGB);
            //绘制改变尺寸后的图
            tag.getGraphics().drawImage(image, 0, 0, expectedWidth, expectedHeight, null);
            //输出流
            BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream("E:/copy.png"));
        }

//        final double scaleW = actualWidth / expectedWidth1;
//        final double scaleH = actualHeight / expectedHeight1;
//        final double scale = NumberUtils.min(scaleW, scaleH);
//
//        final int expectedWidth;
//        final int expectedHeight;
//        if (scale > 1.0f) {
//            expectedWidth = (int) (actualWidth / scale);
//            expectedHeight = (int) (actualHeight / scale);
//        } else {
//            expectedWidth = expectedWidth0;
//            expectedHeight = expectedHeight0;
//        }


//        if (actualHeight > expectedHeight || actualWidth > expectedWidth) {
//            // 需要缩放
//        }

//        int expectedHeight;
//        int expectedWidth = NumberUtils.min(expectedWidth0, A4_CONTENT_WIDTH_DXA);
//        if (expectedWidth == A4_CONTENT_WIDTH_DXA) {
//            // 需要缩放高度
//            expectedHeight = (int) (expectedHeight0 * ((double)expectedWidth0 / A4_CONTENT_WIDTH_DXA));
//            expectedHeight = NumberUtils.min(expectedHeight, A4_CONTENT_HEIGHT_DXA);
//
//            if (expectedHeight == A4_CONTENT_HEIGHT_DXA) {
//                // 如果高度超出，需要重新计算预期宽度
//                expectedWidth = (int) (expectedWidth0 * ((double)expectedHeight / A4_CONTENT_HEIGHT_DXA));
//            }
//        } else {
//            expectedHeight = NumberUtils.min(expectedHeight0, A4_CONTENT_HEIGHT_DXA);
//
//            if (expectedHeight == A4_CONTENT_HEIGHT_DXA) {
//                // 如果高度超出，需要重新计算预期宽度
//                expectedWidth = (int) (expectedWidth0 * ((double)expectedHeight0 / A4_CONTENT_HEIGHT_DXA));
//            }
//        }

//        expectedHeight = PoiUnitTool.pixelToDXA(height);


        double widthPoints = Units.pixelToPoints(width);
        double heightPoints = Units.pixelToPoints(height);

        if (widthPoints > A4_CONTENT_WIDTH_DXA) {
            // 计算缩放比例
            BigDecimal scale = new BigDecimal(A4_CONTENT_WIDTH_DXA).divide(new BigDecimal(widthPoints), 3, BigDecimal.ROUND_HALF_UP);
            widthPoints = new BigDecimal(widthPoints).multiply(scale).setScale(0, BigDecimal.ROUND_HALF_UP).intValue();
            heightPoints = new BigDecimal(heightPoints).multiply(scale).setScale(0, BigDecimal.ROUND_HALF_UP).intValue();
        }

        try (FileInputStream is = new FileInputStream(imgFile)) {
            picture = paragraphRun.addPicture(is, format, imgFile, Units.toEMU(widthPoints), Units.toEMU(heightPoints));
        } catch (InvalidFormatException ignore) {
        }
        return picture;
    }

    /**
     * 设置图片位置
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param leftOffset 水平偏移（单位： 磅）
     * @param topOffset  垂直偏移（单位： 磅）
     */
    public static void setPicturePosition(XWPFParagraph paragraph, int leftOffset, int topOffset) {
        List<XWPFRun> runList = paragraph.getRuns();
        if (runList == null || runList.size() == 0) {
            return;
        }
        XWPFRun paragraphRun = runList.get(runList.size() - 1);
        CTDrawing drawing = paragraphRun.getCTR().getDrawingArray(0);
        CTAnchor ctAnchor = drawing.addNewAnchor();

        ctAnchor.setSimplePos2(false);
        ctAnchor.setRelativeHeight(0);

        // 水平位置
        CTPosH posH;
        if ((posH = ctAnchor.getPositionH()) == null) {
            posH = ctAnchor.addNewPositionH();
        }
        posH.setRelativeFrom(STRelFromH.MARGIN);
        posH.setPosOffset(Units.toEMU(leftOffset));

        // 垂直位置
        CTPosV posV;
        if ((posV = ctAnchor.getPositionV()) == null) {
            posV = ctAnchor.addNewPositionV();
        }
        posV.setRelativeFrom(STRelFromV.PARAGRAPH);
        posV.setPosOffset(Units.toEMU(topOffset));

        // 复制原有的属性
        CTInline ctInline = drawing.getInlineArray(0);

        ctAnchor.setDistT(ctInline.getDistT());
        ctAnchor.setDistR(ctInline.getDistR());
        ctAnchor.setDistB(ctInline.getDistB());
        ctAnchor.setDistL(ctInline.getDistL());

        ctAnchor.setBehindDoc(true);
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

        ctAnchor.addNewWrapNone();
        ctAnchor.addNewCNvGraphicFramePr();

        ctAnchor.setGraphic(ctInline.getGraphic());

        // 移除旧的图片
        drawing.removeInline(0);
    }


    /**
     * 创建没有边框的表格
     *
     * @param document {@link XWPFDocument}
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTableWithoutBorder(XWPFDocument document, int rows, int cols) {
        return createTable(document, rows, cols, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }

    /**
     * 创建表格
     *
     * @param document    {@link XWPFParagraph}
     * @param borderSize  边框宽度
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     * @param rows        行数
     * @param cols        列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTable(XWPFDocument document, int rows, int cols, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        XWPFTable table = document.createTable();
        table.setWidthType(TableWidthType.DXA);
        table.setWidth(String.valueOf(A4_CONTENT_WIDTH_DXA));

        table.setTopBorder(borderType, borderSize, 0, borderColor);
        table.setBottomBorder(borderType, borderSize, 0, borderColor);
        table.setLeftBorder(borderType, borderSize, 0, borderColor);
        table.setRightBorder(borderType, borderSize, 0, borderColor);

        table.setInsideHBorder(borderType, borderSize, 0, borderColor);
        table.setInsideVBorder(borderType, borderSize, 0, borderColor);

        for (int i = 1; i < rows; i++) {
            table.createRow();
        }
        XWPFTableRow tableRowOne = table.getRow(0);

        BigDecimal cellWidth = new BigDecimal(A4_CONTENT_WIDTH_DXA).divide(new BigDecimal(cols), 3, RoundingMode.FLOOR);

        XWPFTableCell tableCell = tableRowOne.getCell(0);
        tableCell.setWidth(String.valueOf(cellWidth.intValue()));
        tableCell.setWidthType(TableWidthType.DXA);

        for (int i = 1; i < cols; i++) {
            tableCell = tableRowOne.addNewTableCell();
            tableCell.setWidth(String.valueOf(cellWidth.intValue()));
            tableCell.setWidthType(TableWidthType.DXA);
        }

        return table;
    }

    /**
     * 设置表格行高
     *
     * @param tableRowOne {@link XWPFTableRow}
     * @param pixel       高度（单位：像素）
     */
    public static void setTableRowHeightOfPixel(XWPFTableRow tableRowOne, int pixel) {
        CTRow ctRow = tableRowOne.getCtRow();
        CTTrPr trPr = (ctRow.isSetTrPr()) ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight ctHeight = trPr.sizeOfTrHeightArray() == 0 ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
        ctHeight.setHRule(STHeightRule.EXACT);
        ctHeight.setVal(BigInteger.valueOf(PoiUnitTool.pixelToDXA(pixel)));
    }

    /**
     * 设置单元格宽度
     *
     * @param tableCell 单元格
     * @param width     宽度（单位：磅）
     */
    public static void setTableWidth(XWPFTableCell tableCell, int width) {
        tableCell.setWidth(String.valueOf(PoiUnitTool.pointToDXA(width)));
        tableCell.setWidthType(TableWidthType.DXA);
    }

    /**
     * 设置单元格文字
     *
     * @param cell            单元格
     * @param horizontalAlign 水平对齐方式
     * @param verticalAlign   垂直对齐方式
     */
    public static void setTableCellAlign(XWPFTableCell cell, String text, String color, STJc.Enum horizontalAlign, XWPFTableCell.XWPFVertAlign verticalAlign) {
        cell.setText(text);
        cell.setColor(color);
        setTableCellAlign(cell, horizontalAlign, verticalAlign);
    }

    /**
     * 设置单元格对齐方式
     *
     * @param cell            单元格
     * @param horizontalAlign 垂直对齐方式
     * @param verticalAlign   水平对齐方式
     */
    public static void setTableCellAlign(XWPFTableCell cell, STJc.Enum horizontalAlign, XWPFTableCell.XWPFVertAlign verticalAlign) {
        if (verticalAlign != null) {
            cell.setVerticalAlignment(verticalAlign);
        }
        // 水平对齐
        if (horizontalAlign != null) {
            CTTc ctTc = cell.getCTTc();
            CTPPr ctpPr;
            List<CTP> ctpList = ctTc.getPList();
            if (ctpList == null || ctpList.size() == 0) {
                return;
            }
            if ((ctpPr = ctpList.get(0).getPPr()) == null) {
                ctpPr = ctpList.get(0).addNewPPr();
            }
            if (ctpPr.isSetJc()) {
                ctpPr.getJc().setVal(horizontalAlign);
            } else {
                ctpPr.addNewJc().setVal(horizontalAlign);
            }
        }

    }

    /**
     * 设置单元格背景色
     *
     * @param cell    单元格
     * @param bgColor 背景色（RGB 格式，例如："FFFFFF"）
     */
    public static void setTableCellBgColor(XWPFTableCell cell, String bgColor) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr cellProperties;
        if ((cellProperties = ctTc.getTcPr()) == null) {
            cellProperties = ctTc.addNewTcPr();
        }
        CTShd ctShd;
        if ((ctShd = cellProperties.getShd()) == null) {
            ctShd = cellProperties.addNewShd();
        }
        ctShd.setColor(bgColor);
        ctShd.setFill(bgColor);
    }
}
