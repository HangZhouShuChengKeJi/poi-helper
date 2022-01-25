package com.orange.poi.word;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

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
     * @param run    段落 {@link XWPFRun}
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
        if (run == null) {
            return;
        }
        run.setText(plainTxt);
        if (StringUtils.isNotBlank(defaultFont)) {
            run.setFontFamily(defaultFont, XWPFRun.FontCharRange.ascii);
            run.setFontFamily(defaultFont, XWPFRun.FontCharRange.cs);
            run.setFontFamily(defaultFont, XWPFRun.FontCharRange.hAnsi);
        }
        if (StringUtils.isNotBlank(eastAsiaFont)) {
            run.setFontFamily(eastAsiaFont, XWPFRun.FontCharRange.eastAsia);
        } else if (StringUtils.isNotBlank(defaultFont)) {
            // 中文使用默认字体
            run.setFontFamily(defaultFont, XWPFRun.FontCharRange.eastAsia);
        }
        if (fontSize != null) {
            run.setFontSize(fontSize);
        }
        if (StringUtils.isNotBlank(color)) {
            run.setColor(color);
        }
        run.setBold(bold);
        if (underline) {
            run.setUnderline(UnderlinePatterns.SINGLE);
        }
    }

    /**
     * 设置上角标
     *
     * @param run  {@link XWPFRun}
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
     * @param run  {@link XWPFRun}
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
     * @param run    段落 {@link XWPFRun}
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
        if (StringUtils.isNotBlank(fontFamily)) {
            run.setFontFamily(fontFamily);
        }
        if (fontSize != null) {
            run.setFontSize(fontSize);
        }
        if (StringUtils.isNotBlank(color)) {
            run.setColor(color);
        }
        if (bold) {
            run.setBold(bold);
        }
        run.setSubscript(verticalAlign);
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
     * 添加制表符
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
