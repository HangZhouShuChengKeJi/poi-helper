package com.orange.poi.word;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

import java.math.BigInteger;

/**
 * 页眉页脚工具
 *
 * @author 小天
 * @date 2022/2/9 11:27
 */
public class PoiWordHeaderFooterTool {

    /**
     * 设置奇偶页页眉页脚不同
     *
     * @param doc               word 文档 {@link XWPFDocument}
     * @param evenAndOddHeaders 是否开启奇偶页不同
     */
    public static void setEvenAndOddHeaders(XWPFDocument doc, boolean evenAndOddHeaders) {
        doc.setEvenAndOddHeadings(evenAndOddHeaders);
    }

    /**
     * 创建页眉
     *
     * @param doc
     * @param sectPr
     *
     * @return
     */
    public static XWPFHeaderFooterPolicy createHeaderFooterPolicy(XWPFDocument doc, CTSectPr sectPr) {
        return new XWPFHeaderFooterPolicy(doc, sectPr);
    }

    /**
     * 创建页眉
     *
     * @param doc
     * @param type
     *
     * @return
     */
    public static XWPFHeader createHeader(XWPFDocument doc, HeaderFooterType type) {
        return createHeader(doc, null, type);
    }

    /**
     * 创建页眉
     *
     * @param doc
     * @param sectPr
     * @param type
     *
     * @return
     */
    public static XWPFHeader createHeader(XWPFDocument doc, CTSectPr sectPr, HeaderFooterType type) {
        XWPFHeaderFooterPolicy hfPolicy = createHeaderFooterPolicy(doc, sectPr);
        if (type == HeaderFooterType.FIRST) {
            CTSectPr ctSectPr = PoiWordSectionTool.getSectionProperties(doc);
            if (!ctSectPr.isSetTitlePg()) {
                CTOnOff titlePg = ctSectPr.addNewTitlePg();
                titlePg.setVal(STOnOff1.ON);
            }
        }
        return createHeader(hfPolicy, type);
    }

    /**
     * 创建页眉
     *
     * @param hfPolicy
     * @param type
     *
     * @return
     */
    public static XWPFHeader createHeader(XWPFHeaderFooterPolicy hfPolicy, HeaderFooterType type) {
        return hfPolicy.createHeader(STHdrFtr.Enum.forInt(type.toInt()));
    }

    /**
     * 创建页脚
     *
     * @param doc
     * @param type
     *
     * @return
     */
    public static XWPFFooter createFooter(XWPFDocument doc, HeaderFooterType type) {
        return createFooter(doc, null, type);
    }

    /**
     * 创建页脚
     *
     * @param doc
     * @param sectPr
     * @param type
     *
     * @return
     */
    public static XWPFFooter createFooter(XWPFDocument doc, CTSectPr sectPr, HeaderFooterType type) {
        XWPFHeaderFooterPolicy hfPolicy = createHeaderFooterPolicy(doc, sectPr);
        if (type == HeaderFooterType.FIRST) {
            CTSectPr ctSectPr = PoiWordSectionTool.getSectionProperties(doc);
            if (!ctSectPr.isSetTitlePg()) {
                CTOnOff titlePg = ctSectPr.addNewTitlePg();
                titlePg.setVal(STOnOff1.ON);
            }
        }
        return createFooter(hfPolicy, type);
    }

    /**
     * 创建页脚
     *
     * @param hfPolicy
     * @param type
     *
     * @return
     */
    public static XWPFFooter createFooter(XWPFHeaderFooterPolicy hfPolicy, HeaderFooterType type) {
        return hfPolicy.createFooter(STHdrFtr.Enum.forInt(type.toInt()));
    }

    /**
     * 设置页眉距离顶部的距离
     *
     * @param doc       word 文档 {@link XWPFDocument}
     * @param marginTop 页眉距离顶部的距离（单位：dxa）
     */
    public static void setHeaderMargin(XWPFDocument doc, long marginTop) {
        CTPageMar pageMar = PoiWordSectionTool.getPageMar(doc);
        pageMar.setHeader(BigInteger.valueOf(marginTop));
    }

    /**
     * 设置页眉距离顶部的距离
     *
     * @param ctSectPr  section 属性 {@link CTSectPr}
     * @param marginTop 页眉距离顶部的距离（单位：dxa）
     */
    public static void setHeaderMargin(CTSectPr ctSectPr, long marginTop) {
        CTPageMar pageMar = PoiWordSectionTool.getPageMar(ctSectPr);
        pageMar.setHeader(BigInteger.valueOf(marginTop));
    }

    /**
     * 设置页脚距离底部的距离
     *
     * @param doc          word 文档 {@link XWPFDocument}
     * @param marginBottom 设置页脚距离底部的距离（单位：dxa）
     */
    public static void setFooterMargin(XWPFDocument doc, long marginBottom) {
        CTPageMar pageMar = PoiWordSectionTool.getPageMar(doc);
        pageMar.setFooter(BigInteger.valueOf(marginBottom));
    }

    /**
     * 设置页脚距离底部的距离
     *
     * @param ctSectPr     section 属性 {@link CTSectPr}
     * @param marginBottom 设置页脚距离底部的距离（单位：dxa）
     */
    public static void setFooterMargin(CTSectPr ctSectPr, long marginBottom) {
        CTPageMar pageMar = PoiWordSectionTool.getPageMar(ctSectPr);
        pageMar.setFooter(BigInteger.valueOf(marginBottom));
    }
}
