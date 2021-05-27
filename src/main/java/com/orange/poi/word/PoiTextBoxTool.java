package com.orange.poi.word;

import com.microsoft.schemas.office.word.CTWrap;
import com.microsoft.schemas.office.word.STWrapType;
import com.microsoft.schemas.vml.CTFill;
import com.microsoft.schemas.vml.CTGroup;
import com.microsoft.schemas.vml.CTShape;
import com.microsoft.schemas.vml.CTShapetype;
import com.microsoft.schemas.vml.CTStroke;
import com.microsoft.schemas.vml.CTTextbox;
import com.microsoft.schemas.vml.STExt;
import com.microsoft.schemas.vml.STFillType;
import com.microsoft.schemas.vml.STStrokeJoinStyle;
import com.orange.poi.PoiUnitTool;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTxbxContent;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHint;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.UUID;

/**
 * 基于 vml 的文本框
 *
 * @author 小天
 * @date 2021/5/18 10:17
 * @see <a href="https://docs.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-introduction">VML Introduction</a>
 */
public class PoiTextBoxTool {

    public final static javax.xml.namespace.QName CTFILL_R_ID_QNAME = new javax.xml.namespace.QName("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id", "r");

    /**
     * 添加文本框到段落中（在文本框处理结束后，再调用）
     *
     * @param xwpfParagraph
     * @param ctGroup
     */
    public static void addGroup(XWPFParagraph xwpfParagraph, CTGroup ctGroup) {
        XWPFRun xwpfRun = xwpfParagraph.createRun();
        CTR ctr = xwpfRun.getCTR();
        CTPicture pict = ctr.addNewPict();
        pict.set(ctGroup);
    }

    /**
     * 创建文本框（默认无边框，默认无填充，默认环绕方式： 上下型）
     *
     * @param width  宽度（单位： pt）
     * @param height 高度（单位： pt）
     *
     * @return
     */
    public static CTGroup createTextBox(double width, double height) {
        return createTextBox(width, height, null, 0, 0);
    }

    /**
     * 创建文本框（默认无边框，默认无填充，默认环绕方式： 上下型）
     *
     * @param width              宽度（单位： pt）
     * @param height             高度（单位： pt）
     * @param positionHorizontal 水平位置（可选值： absolute, left, center, right, inside, outside）
     *
     * @return
     */
    public static CTGroup createTextBox(double width, double height,
                                        String positionHorizontal) {
        return createTextBox(width, height, positionHorizontal, 0, 0);
    }

    /**
     * 创建文本框（默认无边框，默认无填充，默认环绕方式： 上下型）
     *
     * @param width  宽度（单位： pt）
     * @param height 高度（单位： pt）
     * @param left   左偏移（单位： pt）
     * @param top    上偏移（单位： pt）
     *
     * @return
     */
    public static CTGroup createTextBox(double width, double height,
                                        double left, double top) {
        return createTextBox(width, height, null, left, top);
    }

    /**
     * 创建文本框（默认无边框，默认无填充，默认环绕方式： 上下型）
     *
     * @param width              宽度（单位： pt）
     * @param height             高度（单位： pt）
     * @param positionHorizontal 水平位置（可选值： absolute, left, center, right, inside, outside）
     * @param left               左偏移（单位： pt）
     * @param top                上偏移（单位： pt）
     *
     * @return
     */
    public static CTGroup createTextBox(double width, double height,
                                        String positionHorizontal,
                                        double left, double top) {
        // 创建画图板
        CTGroup group = CTGroup.Factory.newInstance();

        // 添加形状类型
        CTShapetype ctShapetype = group.addNewShapetype();
        ctShapetype.setId(StringUtils.remove(UUID.randomUUID().toString(), '-'));
        ctShapetype.setCoordsize((width * 10) + "," + (height * 10));

        // 添加形状
        CTShape ctShape = group.addNewShape();
        ctShape.setId(StringUtils.remove(UUID.randomUUID().toString(), '-'));
        ctShape.setSpid(ctShape.getId());
        ctShape.setType(ctShapetype.getId());

        // 设置 style
        String style = "position:absolute;";
        style += "width:" + width + "pt;height:" + height + "pt;";
        if (StringUtils.isEmpty(positionHorizontal)) {
            style += "left:" + left + "pt;top:" + top + "pt;";
            style += "margin-left:" + left + "pt;margin-top:" + top + "pt;";
        } else {
            style += "left:0pt;top:0pt;";
            style += "margin-left:0pt;margin-top:0pt;";
            style += "mso-position-horizontal:" + positionHorizontal + ";";
        }
        style += "z-index:251660288;";

        style += "mso-wrap-distance-bottom:0pt;mso-wrap-distance-top:0pt;";
        style += "mso-width-relative:page;mso-height-relative:page;";
        ctShape.setStyle(style);

        ctShape.setFillcolor("#FFFFFF");
        ctShape.setSpt(202);
        ctShape.setStroked(STTrueFalse.F);
        ctShape.setFilled(STTrueFalse.F);

        ctShape.addNewPath();
        ctShape.addNewImagedata().setTitle("");
        ctShape.addNewLock().setExt(STExt.EDIT);

        // 边框
        CTStroke ctStroke = ctShape.addNewStroke();
        ctStroke.setOn(STTrueFalse.F);
        ctStroke.setJoinstyle(STStrokeJoinStyle.MITER);

        // 设置文字环绕为： 上下环绕
        CTWrap ctWrap = ctShape.addNewWrap();
        ctWrap.setType(STWrapType.TOP_AND_BOTTOM);

        return group;
    }

    /**
     * 设置背景图
     *
     * @param xwpfDocument
     * @param group
     * @param bgImgFile
     * @param pictureType
     *
     * @throws FileNotFoundException
     * @throws InvalidFormatException
     */
    public static void setBackgroundImg(XWPFDocument xwpfDocument, CTGroup group,
                                        File bgImgFile, int pictureType) throws FileNotFoundException, InvalidFormatException {

        CTShape ctShape = group.getShapeArray(0);

        InputStream pictureData = new FileInputStream(bgImgFile);

        String relationId = xwpfDocument.addPictureData(pictureData, pictureType);

        CTFill ctFill = ctShape.addNewFill();
        ctFill.setType(STFillType.FRAME);
        ctFill.setFocussize("0,0");
        ctFill.setOn(STTrueFalse.T);
        ctFill.setTitle("");
        ctFill.setRecolor(STTrueFalse.T);

        // 通过 XmlCursor 设置 "r:id" 属性
        XmlCursor xmlCursor = ctFill.newCursor();
        xmlCursor.setAttributeText(CTFILL_R_ID_QNAME, relationId);
        xmlCursor.dispose();

    }

    /**
     * 设置文本框内容（不加粗；无下划线；左对齐）
     *
     * @param group        {@link CTGroup}
     * @param plainTxt     文本内容
     * @param asciiFont    ascii 文本字体
     * @param eastAsiaFont 东亚文字字体
     * @param fontSize     字体大小
     * @param color        文本颜色
     */
    public static void setText(CTGroup group, String plainTxt,
                               String asciiFont, String eastAsiaFont, Integer fontSize, String color) {
        setText(group, plainTxt, asciiFont, eastAsiaFont, fontSize, color, false, false, null);
    }

    /**
     * 设置文本框内容
     *
     * @param group        {@link CTGroup}
     * @param plainTxt     文本内容
     * @param asciiFont    ascii 文本字体
     * @param eastAsiaFont 东亚文字字体
     * @param fontSize     字体大小
     * @param color        文本颜色
     * @param bold         是否加粗
     * @param underline    是否下划线
     */
    public static void setText(CTGroup group, String plainTxt,
                               String asciiFont, String eastAsiaFont, Integer fontSize, String color,
                               boolean bold, boolean underline) {
        setText(group, plainTxt, asciiFont, eastAsiaFont, fontSize, color, bold, underline, null);
    }

    /**
     * 设置文本框内容
     *
     * @param group        {@link CTGroup}
     * @param plainTxt     文本内容
     * @param asciiFont    ascii 文本字体
     * @param eastAsiaFont 东亚文字字体
     * @param fontSize     字体大小
     * @param color        文本颜色
     * @param bold         是否加粗
     * @param underline    是否下划线
     * @param stJc         文本水平对齐方式
     */
    public static void setText(CTGroup group, String plainTxt,
                               String asciiFont, String eastAsiaFont, Integer fontSize, String color,
                               boolean bold, boolean underline,
                               STJc.Enum stJc) {
        CTPPr ctpPr = getCTPPr(group);

        if (stJc != null) {
            // 水平对齐方式
            CTJc ctJc;
            if ((ctJc = ctpPr.getJc()) == null) {
                ctJc = ctpPr.addNewJc();
            }
            ctJc.setVal(stJc);
        }

        CTP ctp = getCTP(group);
        addText(ctp, plainTxt, asciiFont, eastAsiaFont, fontSize, color, bold, underline);
    }


    private static void addText(CTP ctp, String plainTxt,
                                String asciiFont, String eastAsiaFont, Integer fontSize, String color,
                                boolean bold, boolean underline) {
        CTR ctr = ctp.addNewR();

        CTRPr ctrPr = ctr.addNewRPr();

        if (bold) {
            ctrPr.addNewB().setVal(true);
        }

        if (underline) {
            ctrPr.addNewU().setVal(STUnderline.SINGLE);
        }

        CTFonts ctFonts = ctrPr.addNewRFonts();
        ctFonts.setHint(STHint.DEFAULT);
        ctFonts.setAscii(asciiFont);
        ctFonts.setEastAsia(eastAsiaFont);

        // 字体大小
        CTHpsMeasure ctSize = ctrPr.addNewSz();
        ctSize.setVal(BigInteger.valueOf(fontSize).multiply(BigInteger.valueOf(2)));

        // 字体颜色
        CTColor ctColor = ctrPr.addNewColor();
        ctColor.setVal(color);

        // 文本内容
        CTText ctText = ctr.addNewT();
        ctText.setStringValue(plainTxt);
    }

    /**
     * 设置段落间距（以磅为单位）
     *
     * @param group      {@link CTGroup}
     * @param lineHeight 行高（单位：磅）
     * @param before     段落前间距（单位：磅）
     * @param after      段落后间距（单位：磅）
     */
    public static void setParagraphSpaceOfPound(CTGroup group, double lineHeight, double before, double after) {
        CTPPr ctpPr = getCTPPr(group);

        CTSpacing spacing;
        if ((spacing = ctpPr.getSpacing()) == null) {
            spacing = ctpPr.addNewSpacing();
        }
        BigInteger line = BigInteger.valueOf(PoiUnitTool.pointToDXA(lineHeight));

        spacing.setLine(line);
        spacing.setLineRule(STLineSpacingRule.EXACT);

        spacing.setBefore(BigInteger.valueOf(PoiUnitTool.pointToDXA(before)));
        spacing.setAfter(BigInteger.valueOf(PoiUnitTool.pointToDXA(after)));
        // 【特别注意】必须同时设置 beforeLines 和 afterLines，比例关系为： 100 / LINE_HEIGHT_DXA
        BigInteger spaceBefore = (BigInteger) spacing.getBefore();
        BigInteger spaceAfter = (BigInteger) spacing.getAfter();
        spacing.setBeforeLines(spaceBefore.divide(line).multiply(BigInteger.valueOf(100)));
        spacing.setAfterLines(spaceAfter.divide(line).multiply(BigInteger.valueOf(100)));
    }

    public static CTP getCTP(CTGroup group) {
        CTShape ctShape = group.getShapeArray(0);
        CTTextbox ctTextbox;
        if (ctShape.sizeOfTextboxArray() > 0) {
            ctTextbox = ctShape.getTextboxArray(0);
        } else {
            ctTextbox = ctShape.addNewTextbox();
        }

        CTTxbxContent ctTxbxContent;
        if ((ctTxbxContent = ctTextbox.getTxbxContent()) == null) {
            ctTxbxContent = ctTextbox.addNewTxbxContent();
        }

        CTP ctp;
        if (ctTxbxContent.sizeOfPArray() > 0) {
            ctp = ctTxbxContent.getPArray(0);
        } else {
            ctp = ctTxbxContent.addNewP();
        }
        return ctp;
    }

    public static CTPPr getCTPPr(CTGroup group) {
        CTP ctp = getCTP(group);
        CTPPr ctpPr;
        if ((ctpPr = ctp.getPPr()) == null) {
            ctpPr = ctp.addNewPPr();
        }
        return ctpPr;
    }

    /**
     * 设置段落缩进
     *
     * @param group          {@link CTGroup}
     * @param firstLineChars 首行缩进字符数量（小于等于 0 时，忽略）
     * @param leftChars      左侧缩进字符数量（小于等于 0 时，忽略）
     * @param rightChars     右侧缩进字符数量（小于等于 0 时，忽略）
     */
    public static void setInd(CTGroup group, double leftChars, double rightChars, double firstLineChars) {
        CTP ctp = getCTP(group);
        PoiWordParagraphTool.setInd(ctp, leftChars, rightChars, firstLineChars);
    }
}
