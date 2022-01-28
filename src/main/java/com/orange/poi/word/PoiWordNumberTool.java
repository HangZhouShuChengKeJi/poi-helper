package com.orange.poi.word;

import com.orange.poi.lowlevel.ParagraphBasePropertyTool;
import com.orange.poi.lowlevel.RunPropertyTool;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrGeneral;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMultiLevelType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import java.math.BigInteger;

/**
 * word 编号工具。
 *
 * @author 小天
 * @date 2022/1/27 11:07
 */
public class PoiWordNumberTool {

    /**
     * 创建编号
     *
     * @param doc           文档对象 {@link XWPFDocument}
     * @param abstractNumId 抽象编号Id
     *
     * @return 编号Id
     */
    public static BigInteger createNumber(XWPFDocument doc, BigInteger abstractNumId) {
        XWPFNumbering numbering = doc.createNumbering();
        return numbering.addNum(abstractNumId);
    }

    /**
     * 创建编号
     *
     * @param doc             文档对象 {@link XWPFDocument}
     * @param xwpfAbstractNum 编号对象 {@link XWPFAbstractNum}
     *
     * @return 编号Id
     */
    public static BigInteger createNumber(XWPFDocument doc, XWPFAbstractNum xwpfAbstractNum) {
        XWPFNumbering numbering = doc.createNumbering();
        return numbering.addNum(getAbstractNumId(xwpfAbstractNum));
    }

    /**
     * 创建单级别抽象编号
     *
     * @param doc 文档对象 {@link XWPFDocument}
     *
     * @return 抽象编号对象 {@link XWPFAbstractNum}
     */
    public static XWPFAbstractNum createAbstractNumOfSingleLevel(XWPFDocument doc) {
        return createAbstractNum(doc, STMultiLevelType.SINGLE_LEVEL);
    }

    /**
     * 创建抽象编号
     *
     * @param doc            文档对象 {@link XWPFDocument}
     * @param multiLevelType 多级编号类型
     *
     * @return 抽象编号对象 {@link XWPFAbstractNum}
     */
    public static XWPFAbstractNum createAbstractNum(XWPFDocument doc, STMultiLevelType.Enum multiLevelType) {
        XWPFNumbering numbering = doc.createNumbering();
        XWPFAbstractNum xwpfAbstractNum = new XWPFAbstractNum(null);
        numbering.addAbstractNum(xwpfAbstractNum);
        xwpfAbstractNum.getAbstractNum().addNewMultiLevelType().setVal(multiLevelType);
        return xwpfAbstractNum;
    }

    /**
     * 获取抽象编号Id
     *
     * @param doc           文档对象 {@link XWPFDocument}
     * @param abstractNumId 抽象编号Id
     *
     * @return 抽象编号对象 {@link XWPFAbstractNum}
     */
    public static XWPFAbstractNum getAbstractNum(XWPFDocument doc, BigInteger abstractNumId) {
        XWPFNumbering numbering = doc.createNumbering();
        return numbering.getAbstractNum(abstractNumId);
    }

    /**
     * 获取抽象编号Id
     *
     * @param xwpfAbstractNum  编号对象 {@link XWPFAbstractNum}
     *
     * @return 抽象编号Id
     */
    public static BigInteger getAbstractNumId(XWPFAbstractNum xwpfAbstractNum) {
        return xwpfAbstractNum.getAbstractNum().getAbstractNumId();
    }

    /**
     * 编号缩进。（word 工具里的 "调整列表缩进"）
     *
     * @param xwpfAbstractNum  编号对象 {@link XWPFAbstractNum}
     * @param level            编号级别
     * @param numLeftChars     编号位置。单位：字符
     * @param textHangingChars 文本缩进。单位：字符
     */
    public static void setInd(XWPFAbstractNum xwpfAbstractNum, BigInteger level, double numLeftChars, double textHangingChars) {
        CTAbstractNum ctAbstractNum = xwpfAbstractNum.getCTAbstractNum();
        CTLvl ctLvl = getLevel(ctAbstractNum, level, true);
        CTPPrGeneral pPr;
        if (ctLvl.isSetPPr()) {
            pPr = ctLvl.getPPr();
        } else {
            pPr = ctLvl.addNewPPr();
        }
        ParagraphBasePropertyTool.setInd(pPr, textHangingChars, -1, -1, numLeftChars);
    }


    /**
     * 编号缩进。（word 工具里的 "调整列表缩进"）
     *
     * @param xwpfAbstractNum 编号对象 {@link XWPFAbstractNum}
     * @param level           编号级别
     * @param numLeft         编号位置。单位：磅
     * @param textHanging     文本缩进。单位：磅
     */
    public static void setIndByPoint(XWPFAbstractNum xwpfAbstractNum, BigInteger level, double numLeft, double textHanging) {
        CTAbstractNum ctAbstractNum = xwpfAbstractNum.getCTAbstractNum();
        CTLvl ctLvl = getLevel(ctAbstractNum, level, true);
        CTPPrGeneral pPr;
        if (ctLvl.isSetPPr()) {
            pPr = ctLvl.getPPr();
        } else {
            pPr = ctLvl.addNewPPr();
        }
        ParagraphBasePropertyTool.setIndByPoint(pPr, textHanging, -1, -1, numLeft);
    }

    /**
     * 设置指定编号级别的样式。编号对齐方式：左对齐。
     *
     * @param xwpfAbstractNum 编号对象 {@link XWPFAbstractNum}
     * @param level           编号级别
     * @param create          未找到指定的级别时，是否创建一个新的
     */
    public static CTLvl getLevel(XWPFAbstractNum xwpfAbstractNum, BigInteger level, boolean create) {
        return getLevel(xwpfAbstractNum.getCTAbstractNum(), level, create);
    }

    /**
     * 设置指定编号级别的样式。编号对齐方式：左对齐。
     *
     * @param ctAbstractNum 编号对象 {@link CTAbstractNum}
     * @param level         编号级别
     * @param create        未找到指定的级别时，是否创建一个新的
     */
    public static CTLvl getLevel(CTAbstractNum ctAbstractNum, BigInteger level, boolean create) {
        for (CTLvl ctLvl : ctAbstractNum.getLvlArray()) {
            if (level.equals(ctLvl.getIlvl())) {
                return ctLvl;
            }
        }
        if (create) {
            CTLvl ctLvl = ctAbstractNum.addNewLvl();
            ctLvl.setIlvl(level);
            return ctLvl;
        }
        return null;
    }

    /**
     * 设置指定编号级别的样式。编号对齐方式：左对齐。
     *
     * @param abstractNum   编号对象 {@link XWPFAbstractNum}
     * @param level         编号级别
     * @param start         起始编号
     * @param numFmt        格式类型
     * @param text          格式文本
     * @param defaultFont   默认字体
     * @param eastAsiaFont  中日韩字体
     * @param fontSize      字号
     * @param color         文字颜色
     */
    public static void setLevel(XWPFAbstractNum abstractNum, BigInteger level,
                                int start, STNumberFormat.Enum numFmt, String text,
                                String defaultFont, String eastAsiaFont, Integer fontSize, String color) {
        setLevel(getLevel(abstractNum, level, true),
                start, numFmt, text, STJc.LEFT,
                defaultFont, eastAsiaFont, fontSize, color,
                false, false, false);
    }

    /**
     * 设置指定编号级别的样式
     *
     * @param abstractNum   编号对象 {@link XWPFAbstractNum}
     * @param level         编号级别
     * @param start         起始编号
     * @param numFmt        格式类型
     * @param text          格式文本
     * @param justification 对齐方式
     * @param defaultFont   默认字体
     * @param eastAsiaFont  中日韩字体
     * @param fontSize      字号
     * @param color         文字颜色
     */
    public static void setLevel(XWPFAbstractNum abstractNum, BigInteger level,
                                int start, STNumberFormat.Enum numFmt, String text, STJc.Enum justification,
                                String defaultFont, String eastAsiaFont, Integer fontSize, String color) {
        setLevel(getLevel(abstractNum, level, true),
                start, numFmt, text, justification,
                defaultFont, eastAsiaFont, fontSize, color,
                false, false, false);
    }

    /**
     * 设置指定编号级别的样式
     *
     * @param ctLvl         编号级别对象 {@link CTLvl}
     * @param start         起始编号
     * @param numFmt        格式类型
     * @param text          格式文本
     * @param justification 对齐方式
     * @param defaultFont   默认字体
     * @param eastAsiaFont  中日韩字体
     * @param fontSize      字号
     * @param color         文字颜色
     * @param bold          是否加粗
     * @param underline     是否下划线
     * @param italics       是否倾斜
     */
    public static void setLevel(CTLvl ctLvl,
                                int start, STNumberFormat.Enum numFmt, String text, STJc.Enum justification,
                                String defaultFont, String eastAsiaFont, Integer fontSize, String color,
                                boolean bold, boolean underline, boolean italics) {
        // 起始编号
        ctLvl.addNewStart().setVal(BigInteger.valueOf(start));
        // 编号格式类型
        ctLvl.addNewNumFmt().setVal(numFmt);
        // 编号格式文本
        ctLvl.addNewLvlText().setVal(text);
        // 编号对齐方式
        ctLvl.addNewLvlJc().setVal(justification);

        // 设置文本样式
        CTRPr ctrPr = ctLvl.addNewRPr();
        RunPropertyTool.set(ctrPr,
                defaultFont, eastAsiaFont, fontSize, color,
                bold, underline, italics);
    }
}
