package com.orange.poi.word;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFactory;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.FtrDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.HdrDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

import java.math.BigInteger;

/**
 * 页眉页脚工具。
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
     * 创建页眉。
     * 该方法创建的页眉不会和任何节挂钩。
     * 如有需要可以使用 {@link #addHeader(XWPFDocument, CTSectPr, HeaderFooterType)} 方法创建页眉，并将页眉关联到节上。
     * 也可以使用 {@link #setHeaderReference(XWPFDocument, CTSectPr, HeaderFooterType, XWPFHeader)} 方法，将页眉关联到节上。
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页眉 {@link XWPFHeader}
     */
    public static XWPFHeader createHeader(XWPFDocument doc) {
        HdrDocument hdrDoc = HdrDocument.Factory.newInstance();

        XWPFRelation relation = XWPFRelation.HEADER;
        int i = PoiWordRelationTool.getRelationIndex(doc, relation);

        XWPFHeader header = (XWPFHeader) doc.createRelationship(relation, XWPFFactory.getInstance(), i);
        header.setXWPFDocument(doc);

        CTHdrFtr hdr = header._getHdrFtr();
        header.setHeaderFooter(hdr);
        hdrDoc.setHdr(hdr);
        return header;
    }

    /**
     * 创建页脚。
     * 该方法创建的页脚不会和任何节挂钩。
     * 如有需要可以使用 {@link #addFooter(XWPFDocument, CTSectPr, HeaderFooterType)} 方法创建页脚，并将页脚关联到节上。
     * 也可以使用 {@link #setFooterReference(XWPFDocument, CTSectPr, HeaderFooterType, XWPFFooter)} 方法，将页脚关联到节上。
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页脚 {@link XWPFFooter}
     */
    public static XWPFFooter createFooter(XWPFDocument doc) {
        FtrDocument ftrDoc = FtrDocument.Factory.newInstance();

        XWPFRelation relation = XWPFRelation.FOOTER;
        int i = PoiWordRelationTool.getRelationIndex(doc, relation);

        XWPFFooter footer = (XWPFFooter) doc.createRelationship(relation, XWPFFactory.getInstance(), i);
        footer.setXWPFDocument(doc);

        CTHdrFtr hdr = footer._getHdrFtr();
        footer.setHeaderFooter(hdr);
        ftrDoc.setFtr(hdr);
        return footer;
    }

    /**
     * 添加页眉到默认的节
     *
     * @param doc  word 文档 {@link XWPFDocument}
     * @param type 页眉类型 {@link HeaderFooterType}
     *
     * @return 页眉 {@link XWPFHeader}
     */
    public static XWPFHeader addHeader(XWPFDocument doc, HeaderFooterType type) {
        return addHeader(doc, PoiWordSectionTool.getSectPr(doc), type);
    }

    /**
     * 添加指定节的页眉
     *
     * @param doc    word 文档 {@link XWPFDocument}
     * @param sectPr 节属性
     * @param type   页眉类型 {@link HeaderFooterType}
     *
     * @return 页眉 {@link XWPFHeader}
     */
    public static XWPFHeader addHeader(XWPFDocument doc, CTSectPr sectPr, HeaderFooterType type) {
        XWPFHeader header = createHeader(doc);
        setHeaderReference(doc, sectPr, type, header);
        return header;
    }

    /**
     * 设置指定节的页眉
     *
     * @param doc              word 文档 {@link XWPFDocument}
     * @param sectPr           节属性 {@link CTSectPr}
     * @param headerFooterType 页眉类型 {@link HeaderFooterType}
     * @param header           页眉 {@link XWPFHeader}
     */
    public static void setHeaderReference(XWPFDocument doc, CTSectPr sectPr, HeaderFooterType headerFooterType, XWPFHeader header) {
        STHdrFtr.Enum type = STHdrFtr.Enum.forInt(headerFooterType.toInt());
        for (CTHdrFtrRef ref : sectPr.getFooterReferenceArray()) {
            if (ref.getType().equals(type)) {
                ref.setId(doc.getRelationId(header));
                return;
            }
        }
        CTHdrFtrRef ref = sectPr.addNewHeaderReference();
        ref.setType(type);
        ref.setId(doc.getRelationId(header));
    }

    /**
     * 添加页脚到默认的节
     *
     * @param doc  word 文档 {@link XWPFDocument}
     * @param type 页脚类型 {@link HeaderFooterType}
     *
     * @return 页脚 {@link XWPFFooter}
     */
    public static XWPFFooter addFooter(XWPFDocument doc, HeaderFooterType type) {
        return addFooter(doc, PoiWordSectionTool.getSectPr(doc), type);
    }

    /**
     * 添加指定节的页脚
     *
     * @param doc    word 文档 {@link XWPFDocument}
     * @param sectPr 节属性 {@link CTSectPr}
     * @param type   页脚类型 {@link HeaderFooterType}
     *
     * @return 页脚 {@link XWPFFooter}
     */
    public static XWPFFooter addFooter(XWPFDocument doc, CTSectPr sectPr, HeaderFooterType type) {
        XWPFFooter header = createFooter(doc);
        setFooterReference(doc, sectPr, type, header);
        return header;
    }

    /**
     * 设置指定节的页脚
     *
     * @param doc              word 文档 {@link XWPFDocument}
     * @param sectPr           节属性 {@link CTSectPr}
     * @param headerFooterType 页脚类型 {@link HeaderFooterType}
     * @param footer           页脚 {@link XWPFFooter}
     */
    public static void setFooterReference(XWPFDocument doc, CTSectPr sectPr, HeaderFooterType headerFooterType, XWPFFooter footer) {
        STHdrFtr.Enum type = STHdrFtr.Enum.forInt(headerFooterType.toInt());
        for (CTHdrFtrRef ref : sectPr.getFooterReferenceArray()) {
            if (ref.getType().equals(type)) {
                ref.setId(doc.getRelationId(footer));
                return;
            }
        }
        CTHdrFtrRef ref = sectPr.addNewFooterReference();
        ref.setType(type);
        ref.setId(doc.getRelationId(footer));
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
