package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.paper.PaperSize;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.officeDocument.x2006.extendedProperties.CTProperties;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STDocGrid;

import java.math.BigInteger;

/**
 * apache poi word 工具类
 *
 * @author 小天
 * @date 2019/5/27 22:23
 * @see <a href="http://poi.apache.org/">Apache POI</a>
 */
public class PoiWordTool {
    /**
     * 基准行高
     */
    public static final int LINE_HEIGHT_DXA = 240;

    /**
     * A4 纸页边距，单位：毫米
     */
    public static final int A4_MARGIN = 15;
    /**
     * B5 纸页边距，单位：毫米
     */
    public static final int B5_MARGIN = 15;


    /**
     * 创建 A4 大小的 word 文档
     *
     * @return 文档 {@link XWPFDocument}
     */
    public static XWPFDocument createDocForA4() {
        XWPFDocument doc = new XWPFDocument();
        initDocForA4(doc);
        return doc;
    }

    /**
     * 创建 B5 大小的 word 文档
     *
     * @return 文档 {@link XWPFDocument}
     */
    public static XWPFDocument createDocForB5() {
        XWPFDocument doc = new XWPFDocument();
        initDocForB5(doc);
        return doc;
    }

    /**
     * 初始化文档尺寸为 A4，页边距：15mm
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @see PaperSize#A4
     */
    public static void initDocForA4(XWPFDocument doc) {
        initDocSize(doc, PaperSize.A4, A4_MARGIN);
    }

    /**
     * 初始化文档尺寸为 B5，页边距：10mm
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @see PaperSize#B5
     */
    public static void initDocForB5(XWPFDocument doc) {
        initDocSize(doc, PaperSize.B5, B5_MARGIN);
    }

    /**
     * 初始化文档尺寸为 A4，页边距：15mm
     *
     * @param doc    文档 {@link XWPFDocument}
     * @param margin 页边距，单位：毫米
     *
     * @see PaperSize#A4
     */
    public static void initDocForA4(XWPFDocument doc, int margin) {
        initDocSize(doc, PaperSize.A4, margin);
    }

    /**
     * 初始化文档尺寸为 B5，页边距：10mm
     *
     * @param doc    文档 {@link XWPFDocument}
     * @param margin 页边距，单位：毫米
     *
     * @see PaperSize#B5
     */
    public static void initDocForB5(XWPFDocument doc, int margin) {
        initDocSize(doc, PaperSize.B5, margin);
    }

//    /**
//     * 初始化文档尺寸
//     *
//     * @param doc    文档 {@link XWPFDocument}
//     * @param width  页面宽度，单位：毫米
//     * @param height 页面高度，单位：毫米
//     * @param margin 页边距，单位：毫米
//     */
//    public static void initDocSize(XWPFDocument doc, int width, int height, int margin) {
//
//        final long widthDxa = PoiUnitTool.centimeterToDXA(width / 10.f);
//        final long heightDxa = PoiUnitTool.centimeterToDXA(height / 10.f);
//
//        final long marginTopDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
//        final long marginBottomDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
//        final long marginLeftDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
//        final long marginRightDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
//
//        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
//        CTPageSz pageSize = sectPr.addNewPgSz();
//        pageSize.setW(BigInteger.valueOf(widthDxa));
//        pageSize.setH(BigInteger.valueOf(heightDxa));
//
//        CTPageMar pageMar = sectPr.addNewPgMar();
//        pageMar.setTop(BigInteger.valueOf(marginTopDxa));
//        pageMar.setRight(BigInteger.valueOf(marginBottomDxa));
//        pageMar.setBottom(BigInteger.valueOf(marginLeftDxa));
//        pageMar.setLeft(BigInteger.valueOf(marginRightDxa));
//
//        /**
//         *  docGrid 的 type 设置为 STDocGrid.LINES，linePitch 设置为 312 被证明在 A4 版面中，用于设置文字在行中居中。
//         *
//         *  todo 这两个值是 通过对比 office 和 wps 生成的文档得出的结论，尚不清楚具体意思
//         */
//        CTDocGrid docGrid = sectPr.addNewDocGrid();
//        docGrid.setType(STDocGrid.LINES);
//        // 线间距，单位：dxa
//        docGrid.setLinePitch(BigInteger.valueOf(312));
//    }

    /**
     * 初始化文档尺寸
     *
     * @param doc       文档 {@link XWPFDocument}
     * @param paperSize 纸张尺寸
     * @param margin    页边距，单位：毫米
     */
    public static void initDocSize(XWPFDocument doc, PaperSize paperSize, int margin) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageSz pageSize = sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(paperSize.width_dxa));
        pageSize.setH(BigInteger.valueOf(paperSize.height_dxa));
        final long marginTopDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
        final long marginBottomDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
        final long marginLeftDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);
        final long marginRightDxa = PoiUnitTool.centimeterToDXA(margin / 10.f);

        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setTop(BigInteger.valueOf(marginTopDxa));
        pageMar.setRight(BigInteger.valueOf(marginBottomDxa));
        pageMar.setBottom(BigInteger.valueOf(marginLeftDxa));
        pageMar.setLeft(BigInteger.valueOf(marginRightDxa));

        /**
         *  docGrid 的 type 设置为 STDocGrid.LINES，linePitch 设置为 312 被证明在 A4 版面中，用于设置文字在行中居中。
         *
         *  todo 这两个值是 通过对比 office 和 wps 生成的文档得出的结论，尚不清楚具体意思
         */
        CTDocGrid docGrid = sectPr.addNewDocGrid();
        docGrid.setType(STDocGrid.LINES);
        // 线间距，单位：dxa
        docGrid.setLinePitch(BigInteger.valueOf(312));
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    public static CTPageSz getPageSize(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageSz pageSize = sectPr.getPgSz();
        if (pageSize == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        return pageSize;
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    public static long getWidthOfDxa(XWPFDocument doc) {
        return getPageSize(doc).getW().longValue();
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    public static long getHeightOfDxa(XWPFDocument doc) {
        return getPageSize(doc).getH().longValue();
    }

    /**
     * 获取页面内容宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面内容宽度
     */
    public static long getContentWidthOfDxa(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageSz pageSize = sectPr.getPgSz();
        if (pageSize == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) {
            throw new IllegalStateException("未设置页边距");
        }
        return pageSize.getW().subtract(pageMar.getLeft()).subtract(pageMar.getRight()).longValue();
    }

    /**
     * 获取页面内容高度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面内容高度
     */
    public static long getContentHeightOfDxa(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageSz pageSize = sectPr.getPgSz();
        if (pageSize == null) {
            throw new IllegalStateException("未设置文档尺寸");
        }
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) {
            throw new IllegalStateException("未设置页边距");
        }
        return pageSize.getH().subtract(pageMar.getTop()).subtract(pageMar.getBottom()).longValue();
    }

    /**
     * 设置文档属性
     *
     * @param doc      文档 {@link XWPFDocument}
     * @param title    文档标题
     * @param author   作者
     * @param company  公司
     * @param category 类别
     */
    public static void setProperties(XWPFDocument doc, String title, String author, String company, String category) {
        POIXMLProperties poixmlProperties = doc.getProperties();
        POIXMLProperties.CoreProperties coreProperties = poixmlProperties.getCoreProperties();
        coreProperties.setTitle(title);
        coreProperties.setCreator(author);
        coreProperties.setCategory(category);

        POIXMLProperties.ExtendedProperties extendedProperties = poixmlProperties.getExtendedProperties();

        CTProperties ctProperties = extendedProperties.getUnderlyingProperties();
        ctProperties.setCompany(company);
    }
}
