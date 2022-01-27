package com.orange.poi.lowlevel;

import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import java.math.BigInteger;

/**
 * @author 小天
 * @date 2022/1/27 16:34
 */
public class RunPropertyTool {

    /**
     * 设置 run 属性
     *
     * @param ctrPr        run 属性 {@link CTRPr}
     * @param defaultFont  默认字体
     * @param eastAsiaFont 中日韩文字字体
     * @param fontSize     字号
     * @param color        颜色
     * @param bold         是否加粗
     * @param underline    是否增加下划线
     */
    public static void set(CTRPr ctrPr,
                           String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                           boolean bold, boolean underline) {
        // 设置字体
        CTFonts font = ctrPr.addNewRFonts();
        font.setAscii(defaultFont);
        font.setEastAsia(eastAsiaFont);
        font.setHAnsi(defaultFont);
        // 设置字体大小
        BigInteger fontSizeVal = BigInteger.valueOf(fontSize).multiply(BigInteger.valueOf(2));
        ctrPr.addNewSz().setVal(fontSizeVal);
        ctrPr.addNewSzCs().setVal(fontSizeVal);
        // 设置颜色
        ctrPr.addNewColor().setVal(color);

        if (bold) {
            // 加粗
            ctrPr.addNewB().setVal(STOnOff1.ON);
        }

        if (underline) {
            // 下划线
            ctrPr.addNewU().setVal(STUnderline.SINGLE);
        }
    }
}
