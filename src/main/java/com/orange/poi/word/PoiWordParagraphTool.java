package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRuby;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRubyContent;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRubyPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STRubyAlign;

import java.math.BigInteger;

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
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, ParagraphAlignment paragraphAlignment, TextAlignment textAlignment) {
        return createParagraph(document, paragraphAlignment, textAlignment, false);
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
     * @param tableCell {@link XWPFTableCell}
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFTableCell tableCell) {
        return createParagraph(tableCell, ParagraphAlignment.LEFT, TextAlignment.CENTER, false);
    }

    /**
     * 创建段落
     *
     * @param tableCell          {@link XWPFTableCell}
     * @param paragraphAlignment 段落对齐方式
     * @param textAlignment      文本对齐方式
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFTableCell tableCell, ParagraphAlignment paragraphAlignment, TextAlignment textAlignment) {
        return createParagraph(tableCell, paragraphAlignment, textAlignment, false);
    }

    /**
     * 创建段落
     *
     * @param tableCell          {@link XWPFTableCell}
     * @param paragraphAlignment 段落对齐方式
     * @param textAlignment      文本对齐方式
     * @param snapToGrid         true: 如果定义了文档网格，则对齐到网格
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFTableCell tableCell, ParagraphAlignment paragraphAlignment, TextAlignment textAlignment, boolean snapToGrid) {
        XWPFParagraph paragraph = tableCell.addParagraph();
        setParagraphAlignment(paragraph, paragraphAlignment);
        setTextAlignment(paragraph, textAlignment);
        setSnapToGrid(paragraph, snapToGrid);
        return paragraph;
    }

    /**
     * 创建段落
     *
     * @param document {@link XWPFDocument}
     * @param plainTxt 文本内容
     *
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph createParagraph(XWPFDocument document, String plainTxt) {
        return createParagraph(document, plainTxt, null, null, null, false, false,
                ParagraphAlignment.LEFT, TextAlignment.CENTER);
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
     * 创建段落
     *
     * @param document           {@link XWPFDocument}
     * @param plainTxt           文本内容
     * @param defaultFont        默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont       东亚文字字体（中日韩文字等）。null 时使用 defaultFont
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
                                                String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                                                boolean bold, boolean underline,
                                                ParagraphAlignment paragraphAlignment, TextAlignment textAlignment,
                                                boolean snapToGrid) {
        XWPFParagraph paragraph = createParagraph(document, paragraphAlignment, textAlignment, snapToGrid);
        addTxt(paragraph, plainTxt, defaultFont, eastAsiaFont, fontSize, color, bold, underline);
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
     * 添加文本内容（没有任何样式）
     *
     * @param paragraph 段落 {@link XWPFParagraph}
     * @param plainTxt  文本内容
     *
     * @return {@link XWPFParagraph}
     */
    public static void addTxt(XWPFParagraph paragraph, String plainTxt) {
        addTxt(paragraph, plainTxt, null, null, null, false, false);
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
        addTxt(paragraph, plainTxt, fontFamily, fontFamily, fontSize, color, bold, underline);
    }


    /**
     * 添加文本内容
     *
     * @param paragraph    段落 {@link XWPFParagraph}
     * @param plainTxt     文本内容
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     字号
     * @param color        颜色（RGB 格式，例如："FFFFFF"）
     * @param bold         是否加粗
     * @param underline    是否增加下划线
     */
    public static void addTxt(XWPFParagraph paragraph, String plainTxt,
                              String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                              boolean bold, boolean underline) {
        if (paragraph == null) {
            return;
        }
        XWPFRun paragraphRun = paragraph.createRun();
        paragraphRun.setText(plainTxt);
        if (StringUtils.isNotBlank(defaultFont)) {
            paragraphRun.setFontFamily(defaultFont, XWPFRun.FontCharRange.ascii);
            paragraphRun.setFontFamily(defaultFont, XWPFRun.FontCharRange.cs);
            paragraphRun.setFontFamily(defaultFont, XWPFRun.FontCharRange.hAnsi);
        }
        if (StringUtils.isNotBlank(eastAsiaFont)) {
            paragraphRun.setFontFamily(eastAsiaFont, XWPFRun.FontCharRange.eastAsia);
        } else if (StringUtils.isNotBlank(defaultFont)) {
            // 中文使用默认字体
            paragraphRun.setFontFamily(defaultFont, XWPFRun.FontCharRange.eastAsia);
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
                ctOnOff.setVal(true);
            } else {
                CTOnOff ctOnOff = ppr.addNewSnapToGrid();
                ctOnOff.setVal(true);
            }
        } else {

            // 不对齐到网格

            if (ppr.isSetSnapToGrid()) {
                CTOnOff ctOnOff = ppr.getSnapToGrid();
                ctOnOff.setVal(false);
            } else {
                CTOnOff ctOnOff = ppr.addNewSnapToGrid();
                ctOnOff.setVal(false);
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
     * 设置段落首行缩进
     *
     * @param paragraph      段落
     * @param firstLineChars 首行缩进字符数量（小于等于 0 时，忽略）
     */
    public static void setInt(XWPFParagraph paragraph, double firstLineChars) {
        setInt(paragraph, firstLineChars, -1, -1);
    }


    /**
     * 设置段落左右两侧缩进
     *
     * @param paragraph  段落
     * @param leftChars  左侧缩进字符数量（小于等于 0 时，忽略）
     * @param rightChars 右侧缩进字符数量（小于等于 0 时，忽略）
     */
    public static void setInt(XWPFParagraph paragraph, double leftChars, double rightChars) {
        setInt(paragraph, -1, leftChars, rightChars);
    }

    /**
     * 设置段落缩进
     *
     * @param paragraph      段落
     * @param firstLineChars 首行缩进字符数量（小于等于 0 时，忽略）
     * @param leftChars      左侧缩进字符数量（小于等于 0 时，忽略）
     * @param rightChars     右侧缩进字符数量（小于等于 0 时，忽略）
     */
    public static void setInt(XWPFParagraph paragraph, double firstLineChars, double leftChars, double rightChars) {
        CTPPr ctpPr = getParagraphProperties(paragraph);
        CTInd ctInd;
        if (ctpPr.isSetInd()) {
            ctInd = ctpPr.getInd();
        } else {
            ctInd = ctpPr.addNewInd();
        }
        if (firstLineChars > 0) {
            ctInd.setFirstLineChars(BigInteger.valueOf((long) (firstLineChars * 100)));
        }
        if (leftChars > 0) {
            ctInd.setLeftChars(BigInteger.valueOf((long) (leftChars * 100)));
        }
        if (rightChars > 0) {
            ctInd.setRightChars(BigInteger.valueOf((long) (rightChars * 100)));
        }
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
        int spaceBefore = ((BigInteger) spacing.getBefore()).intValue();
        spacing.setBeforeLines(BigInteger.valueOf(spaceBefore * 100 / LINE_HEIGHT_DXA));
        spacing.setAfterLines(BigInteger.valueOf(spaceBefore * 100 / LINE_HEIGHT_DXA));
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
     * 添加回车符（不产生新的段落，实际上是添加了一个类型为 {@link BreakType#TEXT_WRAPPING} 的回车符）
     *
     * @param paragraph {@link XWPFParagraph}
     */
    public static void addBreak(XWPFParagraph paragraph) {
        addBreak(paragraph, BreakType.TEXT_WRAPPING);
    }

    /**
     * 添加回车符（不产生新的段落）
     *
     * @param paragraph 段落（{@link XWPFParagraph}）
     * @param breakType 类型（{@link BreakType}）
     */
    public static void addBreak(XWPFParagraph paragraph, BreakType breakType) {
        XWPFRun paragraphRun = getLastXWPFRun(paragraph);
        if (paragraphRun == null) {
            paragraphRun = paragraph.createRun();
        }
        paragraphRun.addBreak(breakType);
    }

    /**
     * 在段落末尾添加分页符（实际上是添加了一个类型为 {@link BreakType#PAGE} 的回车符）
     *
     * @param paragraph 段落（{@link XWPFParagraph}）
     */
    public static void addPageBreak(XWPFParagraph paragraph) {
        addBreak(paragraph, BreakType.PAGE);
    }

    /**
     * 在文档当前位置添加分页符（增加一个新的段落，内容为：一个类型为 {@link BreakType#PAGE} 的回车符）
     *
     * @param document 文档（{@link XWPFDocument}）
     */
    public static void addPageBreak(XWPFDocument document) {
        addBreak(document.createParagraph(), BreakType.PAGE);
    }

    /**
     * @param paragraph {@link XWPFParagraph}
     *
     * @return {@link XWPFRun}
     */
    public static XWPFRun getLastXWPFRun(XWPFParagraph paragraph) {
        XWPFRun paragraphRun = null;
        if (paragraph.getRuns() != null && paragraph.getRuns().size() > 0) {
            paragraphRun = paragraph.getRuns().get(paragraph.getRuns().size() - 1);
        }
        return paragraphRun;
    }

    /**
     * 添加加拼音的文字（针对中文语言）
     *
     * @param paragraph          段落
     * @param baseText           基准文字
     * @param rubyText           拼音文字
     * @param baseTextFontFamily 基准文字字体
     * @param baseTextFontSize   基准文字字体大小（单位：磅）
     * @param baseTextColor      基准文字字体颜色
     * @param rubyTextFontFamily 拼音文字字体
     * @param rubyTextFontSize   拼音文字字体大小（单位：磅）
     * @param rubyTextColor      拼音文字字体颜色
     * @param spaceToBaseText    拼音文字和基准文字的间距（单位：磅）
     */
    public static void addRuby(XWPFParagraph paragraph, String baseText, String rubyText,
                               String baseTextFontFamily, int baseTextFontSize, String baseTextColor,
                               String rubyTextFontFamily, int rubyTextFontSize, String rubyTextColor,
                               int spaceToBaseText) {
        addRuby(paragraph, baseText, rubyText, baseTextFontFamily, baseTextFontSize, baseTextColor, rubyTextFontFamily, rubyTextFontSize, rubyTextColor, spaceToBaseText, "zh-CN");
    }

    /**
     * 添加加拼音的文字
     *
     * @param paragraph          段落
     * @param baseText           基准文字
     * @param rubyText           拼音文字
     * @param baseTextFontFamily 基准文字字体
     * @param baseTextFontSize   基准文字字体大小（单位：磅）
     * @param baseTextColor      基准文字字体颜色
     * @param rubyTextFontFamily 拼音文字字体
     * @param rubyTextFontSize   拼音文字字体大小（单位：磅）
     * @param rubyTextColor      拼音文字字体颜色
     * @param spaceToBaseText    拼音文字和基准文字的间距（单位：磅）
     * @param lang               语言
     */
    public static void addRuby(XWPFParagraph paragraph, String baseText, String rubyText,
                               String baseTextFontFamily, int baseTextFontSize, String baseTextColor,
                               String rubyTextFontFamily, int rubyTextFontSize, String rubyTextColor,
                               int spaceToBaseText, String lang) {

        BigInteger realBaseTextFontSize = BigInteger.valueOf(baseTextFontSize).multiply(new BigInteger("2"));
        BigInteger realRubyTextFontSize = BigInteger.valueOf(rubyTextFontSize).multiply(new BigInteger("2"));

        XWPFRun run = paragraph.createRun();
        CTR ctr = run.getCTR();

        CTRPr ctrPr = ctr.addNewRPr();
        // 设置字体
        CTFonts font = ctrPr.addNewRFonts();
        font.setAscii(baseTextFontFamily);
        font.setEastAsia(baseTextFontFamily);
        font.setHAnsi(baseTextFontFamily);
        // 设置字体大小
        ctrPr.addNewSz().setVal(realBaseTextFontSize);
        ctrPr.addNewSzCs().setVal(realBaseTextFontSize);
        // 设置颜色
        ctrPr.addNewColor().setVal(baseTextColor);


        CTRuby ruby = ctr.addNewRuby();

        CTRubyPr rubyPr = ruby.addNewRubyPr();
        // 拼音以居中的方式显示
        rubyPr.addNewRubyAlign().setVal(STRubyAlign.CENTER);
        // 汉字的字体大小
        rubyPr.addNewHpsBaseText().setVal(realBaseTextFontSize);
        // 拼音的字体大小
        rubyPr.addNewHps().setVal(realRubyTextFontSize);
        // 拼音与汉字的垂直间距
        rubyPr.addNewHpsRaise().setVal(BigInteger.valueOf(baseTextFontSize + spaceToBaseText - 1).multiply(new BigInteger("2")));
        // 语言
        rubyPr.addNewLid().setVal(lang);

        // 汉字
        CTRubyContent rubyBaseContent = ruby.addNewRubyBase();
        CTR rubyBaseCtr = rubyBaseContent.addNewR();
        rubyBaseCtr.addNewT().setStringValue(baseText);
        CTRPr rubyBaseCtrpr = rubyBaseCtr.addNewRPr();
        // 设置字体
        CTFonts rubyBaseFont = rubyBaseCtrpr.addNewRFonts();
        rubyBaseFont.setAscii(baseTextFontFamily);
        rubyBaseFont.setEastAsia(baseTextFontFamily);
        rubyBaseFont.setHAnsi(baseTextFontFamily);
        // 设置字体大小
        rubyBaseCtrpr.addNewSz().setVal(realBaseTextFontSize);
        rubyBaseCtrpr.addNewSzCs().setVal(realBaseTextFontSize);
        // 设置颜色
        rubyBaseCtrpr.addNewColor().setVal(baseTextColor);

        // 拼音
        CTRubyContent rubyRtContent = ruby.addNewRt();
        CTR rubyRtCtr = rubyRtContent.addNewR();
        rubyRtCtr.addNewT().setStringValue(rubyText);
        CTRPr rubyRtCtrpr = rubyRtCtr.addNewRPr();
        // 设置字体
        CTFonts rubyRtFont = rubyRtCtrpr.addNewRFonts();
        rubyRtFont.setAscii(rubyTextFontFamily);
        rubyRtFont.setEastAsia(rubyTextFontFamily);
        rubyRtFont.setHAnsi(rubyTextFontFamily);
        // 设置字体大小
        rubyRtCtrpr.addNewSz().setVal(realRubyTextFontSize);
        // 【注意】这里是汉字的字体大小
        rubyRtCtrpr.addNewSzCs().setVal(realBaseTextFontSize);
        // 设置颜色
        rubyRtCtrpr.addNewColor().setVal(rubyTextColor);
    }
}
