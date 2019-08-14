package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;

import java.math.BigInteger;
import java.util.List;

import static com.orange.poi.word.PoiWordTool.LINE_HEIGHT_DXA;

/**
 * apache poi word 段落工具类
 *
 * @author 小天
 * @date 2019/6/3 19:36
 */
public class PoiWordParagraphTool {

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
     * @param document {@link XWPFDocument}
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document) {
        return createParagraph(document, ParagraphAlignment.LEFT, TextAlignment.CENTER, false);
    }

    /**
     * 创建段落
     *
     * @param document   {@link XWPFDocument}
     * @param snapToGrid true: 如果定义了文档网格，则对齐到网格
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, boolean snapToGrid) {
        return createParagraph(document, ParagraphAlignment.LEFT, TextAlignment.CENTER, snapToGrid);
    }

    /**
     * 创建段落
     *
     * @param document           {@link XWPFDocument}
     * @param paragraphAlignment 段落对齐方式
     * @param textAlignment      文本对齐方式
     * @param snapToGrid         true: 如果定义了文档网格，则对齐到网格
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, ParagraphAlignment paragraphAlignment, TextAlignment textAlignment, boolean snapToGrid) {
        XWPFParagraph paragraph = document.createParagraph();
        setParagraphAlignment(paragraph, paragraphAlignment);
        setTextAlignment(paragraph, textAlignment);
        setSnapToGrid(paragraph, snapToGrid);
        return paragraph;
    }

    /**
     * 创建段落
     *
     * @param document   {@link XWPFDocument}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, String plainTxt,
                                                String fontFamily, Integer fontSize) {
        return createParagraph(document, plainTxt, fontFamily, fontSize, "000000", false, false,
                ParagraphAlignment.LEFT, TextAlignment.CENTER);
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
        return createParagraph(document, plainTxt, fontFamily, fontSize, color, false, false,
                ParagraphAlignment.LEFT, TextAlignment.CENTER, false);
    }

    /**
     * 创建段落
     *
     * @param document           {@link XWPFDocument}
     * @param plainTxt           文本内容
     * @param fontFamily         字体
     * @param fontSize           字号
     * @param color              颜色（RGB 格式，例如："FFFFFF"）
     * @param bold               是否加粗
     * @param underline          是否增加下划线
     * @param paragraphAlignment 段落对齐方式
     * @param textAlignment      文本对齐方式
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, String plainTxt,
                                                String fontFamily, Integer fontSize, String color,
                                                boolean bold, boolean underline,
                                                ParagraphAlignment paragraphAlignment, TextAlignment textAlignment) {
        return createParagraph(document, plainTxt, fontFamily, fontSize, color, bold, underline, paragraphAlignment, textAlignment, false);
    }

    /**
     * 创建段落
     *
     * @param document           {@link XWPFDocument}
     * @param plainTxt           文本内容
     * @param fontFamily         字体
     * @param fontSize           字号
     * @param color              颜色（RGB 格式，例如："FFFFFF"）
     * @param bold               是否加粗
     * @param underline          是否增加下划线
     * @param paragraphAlignment 段落对齐方式
     * @param textAlignment      文本对齐方式
     * @param snapToGrid         true: 如果定义了文档网格，则对齐到网格
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, String plainTxt,
                                                String fontFamily, Integer fontSize, String color,
                                                boolean bold, boolean underline,
                                                ParagraphAlignment paragraphAlignment, TextAlignment textAlignment,
                                                boolean snapToGrid) {
        XWPFParagraph paragraph = createParagraph(document, paragraphAlignment, textAlignment, snapToGrid);
        addTxt(paragraph, plainTxt, fontFamily, fontSize, color, bold, underline);
        return paragraph;
    }

    /**
     * 设置文本对齐方式
     *
     * @param paragraph          段落 {@link XWPFDocument}
     * @param paragraphAlignment 段落对齐方式
     */
    public static void setParagraphAlignment(XWPFParagraph paragraph, ParagraphAlignment paragraphAlignment) {
        paragraph.setAlignment(paragraphAlignment);
    }

    /**
     * 设置文本对齐方式
     *
     * @param paragraph     段落 {@link XWPFDocument}
     * @param textAlignment 文本对齐方式
     */
    public static void setTextAlignment(XWPFParagraph paragraph, TextAlignment textAlignment) {
        paragraph.setVerticalAlignment(textAlignment);
    }

    /**
     * 添加文本内容
     *
     * @param paragraph  段落 {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     *
     * @return {@link XWPFParagraph}
     */
    public static void addTxt(XWPFParagraph paragraph, String plainTxt,
                              String fontFamily, Integer fontSize) {
        addTxt(paragraph, plainTxt, fontFamily, fontSize, "000000", false, false);
    }

    /**
     * 添加文本内容
     *
     * @param paragraph  段落 {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void addTxt(XWPFParagraph paragraph, String plainTxt,
                              String fontFamily, Integer fontSize, String color) {
        addTxt(paragraph, plainTxt, fontFamily, fontSize, color, false, false);
    }

    /**
     * 添加文本内容
     *
     * @param paragraph  段落 {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     * @param bold       是否加粗
     * @param underline  是否增加下划线
     */
    public static void addTxt(XWPFParagraph paragraph, String plainTxt,
                              String fontFamily, Integer fontSize, String color,
                              boolean bold, boolean underline) {
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
    }


    /**
     * 添加下角标
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void addSubscript(XWPFParagraph paragraph, String plainTxt, String fontFamily, Integer fontSize, String color) {
        addSubscript(paragraph, plainTxt, fontFamily, fontSize, color, false, VerticalAlign.SUBSCRIPT);
    }

    /**
     * 添加下角标
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     * @param bold       是否加粗
     */
    public static void addSubscript(XWPFParagraph paragraph, String plainTxt,
                                    String fontFamily, Integer fontSize, String color,
                                    boolean bold) {

        addSubscript(paragraph, plainTxt, fontFamily, fontSize, color, bold, VerticalAlign.SUBSCRIPT);
    }


    /**
     * 添加上角标
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void addSuperscript(XWPFParagraph paragraph, String plainTxt, String fontFamily, Integer fontSize, String color) {
        addSubscript(paragraph, plainTxt, fontFamily, fontSize, color, false, VerticalAlign.SUPERSCRIPT);
    }

    /**
     * 添加上角标
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     * @param bold       是否加粗
     */
    public static void addSuperscript(XWPFParagraph paragraph, String plainTxt,
                                      String fontFamily, Integer fontSize, String color,
                                      boolean bold) {

        addSubscript(paragraph, plainTxt, fontFamily, fontSize, color, bold, VerticalAlign.SUPERSCRIPT);
    }

    /**
     * 添加角标
     *
     * @param paragraph     {@link XWPFParagraph}
     * @param plainTxt      文本内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param color         颜色（RGB 格式，例如："FFFFFF"）
     * @param bold          是否加粗
     * @param verticalAlign 对齐方式
     */
    private static void addSubscript(XWPFParagraph paragraph, String plainTxt,
                                     String fontFamily, Integer fontSize, String color,
                                     boolean bold, VerticalAlign verticalAlign) {
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
        if (bold) {
            paragraphRun.setBold(bold);
        }
        paragraphRun.setSubscript(verticalAlign);
    }

    /**
     * 设置行高（行距）为多倍行距
     *
     * @param paragraph {@link XWPFParagraph}
     * @param multiple  多倍行距，例如： 1.5f 表示 1.5 倍行距
     */
    public static void setLineHeightMultiple(XWPFParagraph paragraph, double multiple) {
        CTPPr ppr = getParagraphProperties(paragraph);
        CTSpacing spacing;
        if ((spacing = ppr.getSpacing()) == null) {
            spacing = ppr.addNewSpacing();
        }
        spacing.setLine(BigInteger.valueOf((long) (multiple * LINE_HEIGHT_DXA)));
        spacing.setLineRule(STLineSpacingRule.AUTO);
    }

    /**
     * 设置行高（行距）为固定值。<br>
     * <p><b>注意：</b> 设置行高为固定值时，在 wps 里，文本不能垂直居中。</p>
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
        spacing.setLine(BigInteger.valueOf(PoiUnitTool.pointToDXA(value)));
        spacing.setLineRule(STLineSpacingRule.EXACT);
    }

    /**
     * 设置段落是否对齐到网格。<br>
     *
     * @param paragraph  {@link XWPFParagraph}
     * @param snapToGrid true: 如果定义了文档网格，则对齐到网格
     */
    public static void setSnapToGrid(XWPFParagraph paragraph, boolean snapToGrid) {
        CTPPr ppr = getParagraphProperties(paragraph);
        if (snapToGrid) {

            // 对齐到网格

            if (ppr.isSetSnapToGrid()) {
                CTOnOff ctOnOff = ppr.getSnapToGrid();
                ctOnOff.setVal(STOnOff.X_1);
            } else {
                CTOnOff ctOnOff = ppr.addNewSnapToGrid();
                ctOnOff.setVal(STOnOff.X_1);
            }
        } else {

            // 不对齐到网格

            if (ppr.isSetSnapToGrid()) {
                CTOnOff ctOnOff = ppr.getSnapToGrid();
                ctOnOff.setVal(STOnOff.X_0);
            } else {
                CTOnOff ctOnOff = ppr.addNewSnapToGrid();
                ctOnOff.setVal(STOnOff.X_0);
            }
        }
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
        spacing.setBefore(BigInteger.valueOf(PoiUnitTool.pointToDXA(before)));
        spacing.setAfter(BigInteger.valueOf(PoiUnitTool.pointToDXA(after)));
        // 【特别注意】必须同时设置 beforeLines 和 afterLines，比例关系为： 100 / LINE_HEIGHT_DXA
        spacing.setBeforeLines(BigInteger.valueOf((long) ((double) spacing.getBefore().intValue() * 100 / LINE_HEIGHT_DXA)));
        spacing.setAfterLines(BigInteger.valueOf((long) ((double) spacing.getBefore().intValue() * 100 / LINE_HEIGHT_DXA)));
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
        spacing.setAfter(BigInteger.valueOf((long) (afterLines * LINE_HEIGHT_DXA)));
        // 【特别注意】必须同时设置 beforeLines 和 afterLines，行高的基数为 100
        spacing.setBeforeLines(BigInteger.valueOf((long) (beforeLines * 100)));
        spacing.setAfterLines(BigInteger.valueOf((long) (afterLines * 100)));
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
        List<XWPFRun> runList;
        if (CollectionUtils.isNotEmpty(runList = paragraph.getRuns())) {
            paragraphRun = paragraph.getRuns().get(runList.size() - 1);
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
}
