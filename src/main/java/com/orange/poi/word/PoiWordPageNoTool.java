package com.orange.poi.word;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * 页码工具类
 *
 * @author 小天
 * @date 2022/2/9 11:59
 */
public class PoiWordPageNoTool {

    /**
     * 插入页码
     *
     * @param paragraph  段落 {@link XWPFParagraph}
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     字号
     * @param color        颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void addPageNo(XWPFParagraph paragraph, String defaultFont, String eastAsiaFont, Integer fontSize, String color) {
        PoiWordParagraphTool.addFieldCode(paragraph, "PAGE", defaultFont, eastAsiaFont, fontSize, color);
    }
}
