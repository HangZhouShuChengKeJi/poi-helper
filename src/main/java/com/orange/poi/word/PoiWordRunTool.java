package com.orange.poi.word;

import com.orange.poi.lowlevel.RunPropertyTool;
import com.orange.poi.lowlevel.RunTool;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;

/**
 * run 元素工具。对应 "r" 标签。
 *
 * @author 小天
 * @date 2022/1/25 11:47
 */
public class PoiWordRunTool {

    /**
     * 设置文本内容
     *
     * @param run          段落 {@link XWPFRun}
     * @param plainTxt     文本内容
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     字号
     * @param color        颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void setTxt(XWPFRun run, String plainTxt,
                              String defaultFont, String eastAsiaFont, Integer fontSize, String color) {
        setTxt(run, plainTxt,
                defaultFont, eastAsiaFont, fontSize, color,
                false, false, false);
    }

    /**
     * 设置文本内容
     *
     * @param run          段落 {@link XWPFRun}
     * @param plainTxt     文本内容
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     字号
     * @param color        颜色（RGB 格式，例如："FFFFFF"）
     * @param bold         是否加粗
     * @param underline    是否增加下划线
     */
    public static void setTxt(XWPFRun run, String plainTxt,
                              String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                              boolean bold, boolean underline) {
        setTxt(run, plainTxt, defaultFont, eastAsiaFont, fontSize, color, bold, underline, false);
    }

    /**
     * 设置文本内容
     *
     * @param run          段落 {@link XWPFRun}
     * @param plainTxt     文本内容
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     字号
     * @param color        颜色（RGB 格式，例如："FFFFFF"）
     * @param bold         是否加粗
     * @param underline    是否增加下划线
     * @param italics      是否倾斜
     */
    public static void setTxt(XWPFRun run, String plainTxt,
                              String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                              boolean bold, boolean underline, boolean italics) {
        if (run == null) {
            return;
        }
        run.setText(plainTxt);

        CTR ctr = run.getCTR();
        RunPropertyTool.set(ctr,
                defaultFont, eastAsiaFont, fontSize, color,
                bold, underline, italics);
    }

    /**
     * 设置 instrText（域代码）内容
     *
     * @param run          段落 {@link XWPFRun}
     * @param fieldCode    域代码。
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     字号
     * @param color        颜色（RGB 格式，例如："FFFFFF"）
     * @param bold         是否加粗
     * @param underline    是否增加下划线
     * @param italics      是否倾斜
     */
    public static void setInstrTxt(XWPFRun run, String fieldCode,
                              String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                              boolean bold, boolean underline, boolean italics) {
        if (run == null) {
            return;
        }
        CTR ctr = run.getCTR();

        CTText ctText = ctr.addNewInstrText();
        ctText.setStringValue(fieldCode);
        ctText.setSpace(SpaceAttribute.Space.PRESERVE);

        RunPropertyTool.set(ctr,
                defaultFont, eastAsiaFont, fontSize, color,
                bold, underline, italics);
    }

    /**
     * 获取 instrText（域代码）内容
     *
     * @param run 段落 {@link XWPFRun}
     *
     * @return instrText（域代码）内容
     */
    public static String getInstrTxt(XWPFRun run) {
        if (run == null) {
            return null;
        }
        CTR ctr = run.getCTR();
        return RunTool.getInstrTxt(ctr);
    }

    /**
     * 获取复杂域字符类型。复杂域字符类型是一种特殊字符，有以下几个值：
     * <ul>
     *     <li>{@link STFldCharType#BEGIN}: 复杂域开始</li>
     *     <li>{@link STFldCharType#SEPARATE}: 复杂域分隔符</li>
     *     <li>{@link STFldCharType#END}: 复杂域结束</li>
     * </ul>
     *
     * @param run 段落 {@link XWPFRun}
     *
     * @return 域代码类型 {@link STFldCharType.Enum}
     */
    public static STFldCharType.Enum getFieldCharType(XWPFRun run) {
        if (run == null) {
            return null;
        }
        CTR ctr = run.getCTR();
        return RunTool.getFldCharType(ctr);
    }

    /**
     * 是否为域代码开始标记
     *
     * @param run 段落 {@link XWPFRun}
     *
     * @return
     * @see {@link #getFieldCharType(XWPFRun)}
     */
    public static boolean isFieldBegin(XWPFRun run) {
        STFldCharType.Enum type = getFieldCharType(run);
        return type == STFldCharType.BEGIN;
    }

    /**
     * 是否为域代码分隔符
     *
     * @param run 段落 {@link XWPFRun}
     *
     * @return
     * @see {@link #getFieldCharType(XWPFRun)}
     */
    public static boolean isFieldSeparate(XWPFRun run) {
        STFldCharType.Enum type = getFieldCharType(run);
        return type == STFldCharType.SEPARATE;
    }

    /**
     * 是否为域代码结束标记
     *
     * @param run 段落 {@link XWPFRun}
     *
     * @return
     * @see {@link #getFieldCharType(XWPFRun)}
     */
    public static boolean isFieldEnd(XWPFRun run) {
        STFldCharType.Enum type = getFieldCharType(run);
        return type == STFldCharType.END;
    }

    /**
     * 设置上角标
     *
     * @param run        {@link XWPFRun}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     * @param bold       是否加粗
     */
    public static void setSuperscript(XWPFRun run, String plainTxt,
                                      String fontFamily, Integer fontSize, String color,
                                      boolean bold) {
        setSubscript(run, plainTxt, fontFamily, fontSize, color, bold, VerticalAlign.SUPERSCRIPT);
    }

    /**
     * 设置下角标
     *
     * @param run        {@link XWPFRun}
     * @param plainTxt   文本内容
     * @param fontFamily 字体
     * @param fontSize   字号
     * @param color      颜色（RGB 格式，例如："FFFFFF"）
     * @param bold       是否加粗
     */
    public static void setSubscript(XWPFRun run, String plainTxt,
                                      String fontFamily, Integer fontSize, String color,
                                      boolean bold) {
        setSubscript(run, plainTxt, fontFamily, fontSize, color, bold, VerticalAlign.SUBSCRIPT);
    }

    /**
     * 设置角标
     *
     * @param run           段落 {@link XWPFRun}
     * @param plainTxt      文本内容
     * @param fontFamily    字体
     * @param fontSize      字号
     * @param color         颜色（RGB 格式，例如："FFFFFF"）
     * @param bold          是否加粗
     * @param verticalAlign 对齐方式
     */
    private static void setSubscript(XWPFRun run, String plainTxt,
                                     String fontFamily, Integer fontSize, String color,
                                     boolean bold, VerticalAlign verticalAlign) {
        if (run == null) {
            return;
        }
        run.setText(plainTxt);
        run.setSubscript(verticalAlign);
        RunPropertyTool.set(run.getCTR(),
                fontFamily, fontFamily, fontSize, color,
                bold, false, false);
    }

    /**
     * 添加制表符
     *
     * @param xwpfRun {@link XWPFRun}
     */
    public static void setTab(XWPFRun xwpfRun) {
        CTR ctr = xwpfRun.getCTR();
        if (ctr.sizeOfTabArray() == 0) {
            xwpfRun.addTab();
        }
    }
    /**
     * 删除制表符
     *
     * @param xwpfRun {@link XWPFRun}
     */
    public static void removeTab(XWPFRun xwpfRun) {
        CTR ctr = xwpfRun.getCTR();
        if (ctr.sizeOfTabArray() > 0) {
            ctr.removeTab(0);
        }
    }

}
