package com.orange.poi.word;

import com.orange.poi.util.OfficeMathMLUtil;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.dom4j.DocumentException;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTF;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import javax.xml.transform.TransformerException;

/**
 * apache poi word 数学公式工具类
 *
 * @author 小天
 * @date 2019/7/5 23:01
 */
public class PoiWordMathTool {

    public static void addMathML(XWPFParagraph paragraph, String mathML) throws XmlException, TransformerException, DocumentException {
        // 转换为 office mathml
        String officeMathML = OfficeMathMLUtil.getInstance().convertMmlToOmml(mathML);
        XmlToken xmlToken = XmlToken.Factory.parse(officeMathML, org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS);

        CTOMathPara ctoMathPara = paragraph.getCTP().addNewOMathPara();
        ctoMathPara.set(xmlToken);
    }


    /**
     * 添加分数
     *
     * @param paragraph   段落 {@link XWPFParagraph}
     * @param numerator   分子
     * @param denominator 分母
     */
    public static void addFraction(XWPFParagraph paragraph, String numerator, String denominator) {
        addFraction(paragraph, numerator, denominator, "Cambria Math", "000000");
    }

    /**
     * 添加分数
     *
     * @param paragraph   段落 {@link XWPFParagraph}
     * @param numerator   分子
     * @param denominator 分母
     * @param fontFamily  字体
     * @param color       颜色
     */
    public static void addFraction(XWPFParagraph paragraph, String numerator, String denominator, String fontFamily, String color) {
        CTOMath ctoMath = paragraph.getCTP().addNewOMath();
        CTF ctf = ctoMath.addNewF();

        CTR numCTR = ctf.addNewNum().addNewR();
        numCTR.addNewT().setStringValue(numerator);

        CTR denCTR = ctf.addNewDen().addNewR();
        denCTR.addNewT().setStringValue(denominator);

        CTRPr ctrPr = null;
        CTRPr numCTRPr = null;
        CTRPr denCTRPr = null;
        if (fontFamily != null) {

            ctrPr = ctf.addNewFPr().addNewCtrlPr().addNewRPr();
            CTFonts fonts = ctrPr.addNewRFonts();
            fonts.setAscii(fontFamily);
            fonts.setHAnsi(fontFamily);

            numCTRPr = numCTR.addNewRPr2();
            CTFonts numFont = numCTRPr.addNewRFonts();
            numFont.setAscii(fontFamily);
            numFont.setHAnsi(fontFamily);

            denCTRPr = denCTR.addNewRPr2();
            CTFonts denFont = denCTRPr.addNewRFonts();
            denFont.setAscii(fontFamily);
            denFont.setHAnsi(fontFamily);
        }

        if (color != null) {
            if (ctrPr == null) {
                ctrPr = ctf.addNewFPr().addNewCtrlPr().addNewRPr();
            }
            ctrPr.addNewColor().setVal(color);

            if (numCTRPr == null) {
                numCTRPr = numCTR.addNewRPr2();
            }
            numCTRPr.addNewColor().setVal(color);

            if (denCTRPr == null) {
                denCTRPr = denCTR.addNewRPr2();
            }
            denCTRPr.addNewColor().setVal(color);
        }
    }
}
