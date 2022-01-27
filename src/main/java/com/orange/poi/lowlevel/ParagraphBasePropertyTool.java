package com.orange.poi.lowlevel;

import com.orange.poi.PoiUnitTool;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrBase;

import java.math.BigInteger;

/**
 * 段落基础属性操作工具
 *
 * @author 小天
 * @date 2022/1/27 15:41
 */
public class ParagraphBasePropertyTool {

    /**
     * 设置段落缩进
     *
     * @param ctpPrBase      段落基础属性 {@link CTPPrBase}
     * @param leftChars      左侧缩进（小于等于 0 时，忽略）。单位：字符
     * @param rightChars     右侧缩进（小于等于 0 时，忽略）。单位：字符
     * @param firstLineChars 首行缩进（小于等于 0 时，忽略）。单位：字符
     * @param hangingChars   悬挂缩进（小于等于 0 时，忽略）。单位：字符
     */
    public static void setInd(CTPPrBase ctpPrBase, double leftChars, double rightChars, double firstLineChars, double hangingChars) {
        CTInd ctInd;
        if (ctpPrBase.isSetInd()) {
            ctInd = ctpPrBase.getInd();
        } else {
            ctInd = ctpPrBase.addNewInd();
        }
        setInd(ctInd, leftChars, rightChars, firstLineChars, hangingChars);
    }

    /**
     * 以 "字符" 为单位，设置段落缩进
     *
     * @param ctInd          缩进属性 {@link CTInd}
     * @param leftChars      左侧缩进（小于等于 0 时，忽略）。单位：字符
     * @param rightChars     右侧缩进（小于等于 0 时，忽略）。单位：字符
     * @param firstLineChars 首行缩进（小于等于 0 时，忽略）。单位：字符
     * @param hangingChars   悬挂缩进（小于等于 0 时，忽略）。单位：字符
     */
    public static void setInd(CTInd ctInd, double leftChars, double rightChars, double firstLineChars, double hangingChars) {
        if (leftChars > 0) {
            ctInd.setLeftChars(BigInteger.valueOf((long) (leftChars * 100)));
//            ctInd.setStartChars(BigInteger.valueOf((long) (leftChars * 100)));
        } else {
            if (ctInd.isSetLeftChars()) {
                ctInd.unsetLeftChars();
            }
        }
        if (rightChars > 0) {
            ctInd.setRightChars(BigInteger.valueOf((long) (rightChars * 100)));
//            ctInd.setEndChars(BigInteger.valueOf((long) (rightChars * 100)));
        } else {
            if (ctInd.isSetRightChars()) {
                ctInd.unsetRightChars();
            }
        }
        if (firstLineChars > 0) {
            ctInd.setFirstLineChars(BigInteger.valueOf((long) (firstLineChars * 100)));
        } else {
            if (ctInd.isSetFirstLineChars()) {
                ctInd.unsetFirstLineChars();
            }
        }
        if (hangingChars > 0) {
            ctInd.setHangingChars(BigInteger.valueOf((long) (hangingChars * 100)));
        } else {
            if (ctInd.isSetHangingChars()) {
                ctInd.unsetHangingChars();
            }
        }
    }

    /**
     * 以 "磅" 为单位，设置段落缩进
     *
     * @param ctpPrBase 段落基础属性 {@link CTPPrBase}
     * @param left      左侧缩进（小于等于 0 时，忽略）。单位：磅
     * @param right     右侧缩进（小于等于 0 时，忽略）。单位：磅
     * @param firstLine 首行缩进（小于等于 0 时，忽略）。单位：磅
     * @param hanging   悬挂缩进（小于等于 0 时，忽略）。单位：磅
     */
    public static void setIndByPoint(CTPPrBase ctpPrBase, double left, double right, double firstLine, double hanging) {
        CTInd ctInd;
        if (ctpPrBase.isSetInd()) {
            ctInd = ctpPrBase.getInd();
        } else {
            ctInd = ctpPrBase.addNewInd();
        }
        setIndByPoint(ctInd, left, right, firstLine, hanging);
    }

    /**
     * 以 "磅" 为单位，设置段落缩进
     *
     * @param ctInd     缩进属性 {@link CTInd}
     * @param left      左侧缩进（小于等于 0 时，忽略）。单位：磅
     * @param right     右侧缩进（小于等于 0 时，忽略）。单位：磅
     * @param firstLine 首行缩进（小于等于 0 时，忽略）。单位：磅
     * @param hanging   悬挂缩进（小于等于 0 时，忽略）。单位：磅
     */
    public static void setIndByPoint(CTInd ctInd, double left, double right, double firstLine, double hanging) {
        if (left > 0) {
            ctInd.setLeft(PoiUnitTool.pointToDXA(left));
        } else {
            if (ctInd.isSetLeft()) {
                ctInd.unsetLeft();
            }
        }
        if (right > 0) {
            ctInd.setRight(PoiUnitTool.pointToDXA(right));
        } else {
            if (ctInd.isSetRight()) {
                ctInd.unsetRight();
            }
        }
        if (firstLine > 0) {
            ctInd.setFirstLine(PoiUnitTool.pointToDXA(firstLine));
        } else {
            if (ctInd.isSetFirstLine()) {
                ctInd.unsetFirstLine();
            }
        }
        if (hanging > 0) {
            ctInd.setHanging(PoiUnitTool.pointToDXA(hanging));
        } else {
            if (ctInd.isSetHanging()) {
                ctInd.unsetHanging();
            }
        }
    }
}
