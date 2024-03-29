package com.orange.poi.word;

import com.orange.poi.paper.PaperSize;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.officeDocument.x2006.extendedProperties.CTProperties;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocDefaults;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPrDefault;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STDocGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.StylesDocument;

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
    public static final int A4_MARGIN = 12;
    /**
     * B5 纸页边距，单位：毫米
     */
    public static final int B5_MARGIN = 12;

    /**
     * 创建 word 文档
     *
     * @return 文档 {@link XWPFDocument}
     */
    public static XWPFDocument createDoc() {
        return new XWPFDocument();
    }


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
     * 初始化文档尺寸为 A3.
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @see PaperSize#A4
     */
    public static void initDocForA3(XWPFDocument doc) {
        initDocSize(doc, PaperSize.A3, A4_MARGIN);
    }

    /**
     * 初始化文档尺寸为 A3 横排版式。
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @see PaperSize#A4
     */
    public static void initDocForA3H(XWPFDocument doc) {
        initDocSize(doc, PaperSize.A3_H, A4_MARGIN);
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

    /**
     * 初始化文档尺寸
     *
     * @param doc       文档 {@link XWPFDocument}
     * @param paperSize 纸张尺寸
     * @param margin    页边距，单位：毫米
     */
    public static void initDocSize(XWPFDocument doc, PaperSize paperSize, int margin) {
        CTSectPr ctSectPr = PoiWordSectionTool.getSectPr(doc);
        PoiWordSectionTool.setPageSize(ctSectPr, paperSize, margin, margin, margin, margin);

        /**
         *  docGrid 的 type 设置为 STDocGrid.LINES，linePitch 设置为 312 被证明在 A4 版面中，用于设置文字在行中居中。
         *
         *  todo 这两个值是 通过对比 office 和 wps 生成的文档得出的结论，尚不清楚具体意思
         */
        CTDocGrid docGrid = ctSectPr.addNewDocGrid();
        docGrid.setType(STDocGrid.LINES);
        // 线间距，单位：dxa
        docGrid.setLinePitch(BigInteger.valueOf(312));
    }

    /**
     * 获取 word 默认节
     *
     * @param doc             文档 ({@link XWPFDocument})
     *
     * @return 节对象 {@link CTSectPr}
     *
     * @deprecated 请使用 {@link PoiWordSectionTool#getSectPr(XWPFDocument)} 替代。
     */
    @Deprecated
    public static CTSectPr getSectPr(XWPFDocument doc) {
        return PoiWordSectionTool.getSectPr(doc);
    }

    /**
     * 获取 word 默认节的边距设置。
     *
     * @param doc 文档 ({@link XWPFDocument})
     *
     * @return 边距对象 {@link CTPageMar}
     *
     * @deprecated 请使用 {@link PoiWordSectionTool#getSectPr(XWPFDocument)} 获取默认的节对象，然后使用 {@link PoiWordSectionTool#getPageMar(CTSectPr)} 获取。
     */
    @Deprecated
    public static CTPageMar getPageMar(XWPFDocument doc) {
        CTSectPr sectPr = PoiWordSectionTool.getSectPr(doc);
        return PoiWordSectionTool.getPageMar(sectPr);
    }

    /**
     * 设置页边距
     *
     * @param doc             文档 ({@link XWPFDocument})
     * @param marginTopDxa    上边距（单位：dxa）
     * @param marginBottomDxa 下边距（单位：dxa）
     * @param marginLeftDxa   左边距（单位：dxa）
     * @param marginRightDxa  右边距（单位：dxa）
     *
     * @deprecated 请使用 {@link PoiWordSectionTool#getSectPr(XWPFDocument)} 获取默认的节对象，
     * 然后使用 {@link PoiWordSectionTool#setPageMargin(CTSectPr, long, long, long, long)} 进行设置。
     */
    @Deprecated
    public static void setPageMargin(XWPFDocument doc,
                                     long marginTopDxa, long marginBottomDxa,
                                     long marginLeftDxa, long marginRightDxa) {
        CTSectPr sectPr = PoiWordSectionTool.getSectPr(doc);
        PoiWordSectionTool.setPageMargin(sectPr, marginTopDxa, marginBottomDxa, marginLeftDxa, marginRightDxa);
    }

    /**
     * 设置页眉距离顶部的距离
     *
     * @param doc       文档 ({@link XWPFDocument})
     * @param marginTop 页眉距离顶部的距离（单位：dxa）
     *
     * @deprecated 请使用 {@link PoiWordHeaderFooterTool#setHeaderMargin(XWPFDocument, long)} 替代。
     */
    @Deprecated
    public static void setHeaderMargin(XWPFDocument doc, long marginTop) {
        PoiWordHeaderFooterTool.setHeaderMargin(doc, marginTop);
    }

    /**
     * 设置页脚距离底部的距离
     *
     * @param doc          文档 ({@link XWPFDocument})
     * @param marginBottom 设置页脚距离底部的距离（单位：dxa）
     *
     * @deprecated 请使用 {@link PoiWordHeaderFooterTool#setFooterMargin(XWPFDocument, long)} 替代。
     */
    @Deprecated
    public static void setFooterMargin(XWPFDocument doc, long marginBottom) {
        PoiWordHeaderFooterTool.setFooterMargin(doc, marginBottom);
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    public static CTPageSz getPageSize(XWPFDocument doc) {
        CTSectPr sectPr = PoiWordSectionTool.getSectPr(doc);
        return PoiWordSectionTool.getPageSize(sectPr);
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    public static long getWidthOfDxa(XWPFDocument doc) {
        return ((BigInteger)getPageSize(doc).getW()).longValue();
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    public static long getHeightOfDxa(XWPFDocument doc) {
        return ((BigInteger)getPageSize(doc).getH()).longValue();
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
        return ((BigInteger)pageSize.getW()).subtract((BigInteger)pageMar.getLeft()).subtract((BigInteger)pageMar.getRight()).longValue();
    }

    /**
     * 获取页面内容宽度，单位：dxa
     *
     * @param pageSize {@link CTPageSz}
     * @param pageMar {@link CTPageMar}
     *
     * @return 页面内容宽度
     */
    public static long getContentWidthOfDxa(CTPageSz pageSize, CTPageMar pageMar) {
        return ((BigInteger)pageSize.getW()).subtract((BigInteger)pageMar.getLeft()).subtract((BigInteger)pageMar.getRight()).longValue();
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
        return ((BigInteger)pageSize.getH()).subtract((BigInteger)pageMar.getTop()).subtract((BigInteger)pageMar.getBottom()).longValue();
    }


    /**
     * 获取页面内容高度，单位：dxa
     *
     * @param pageSize {@link CTPageSz}
     * @param pageMar {@link CTPageMar}
     *
     * @return 页面内容高度
     */
    public static long getContentHeightOfDxa(CTPageSz pageSize, CTPageMar pageMar) {
        return ((BigInteger)pageSize.getH()).subtract((BigInteger)pageMar.getTop()).subtract((BigInteger)pageMar.getBottom()).longValue();
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

    /**
     * 设置文档默认样式
     *
     * @param doc          文档 {@link XWPFDocument}
     * @param defaultFont  默认字体（用于 ascii 等字符的字体）
     * @param eastAsiaFont 东亚文字字体（中日韩文字等）。null 时使用 defaultFont
     * @param fontSize     默认字体大小
     * @param color        文字颜色
     */
    public static CTStyles setDefaultStyle(XWPFDocument doc, String defaultFont, String eastAsiaFont, Integer fontSize, String color) {
        StylesDocument stylesDoc = (StylesDocument) StylesDocument.Factory.newInstance();
        CTStyles ctStyles = stylesDoc.addNewStyles();
        CTDocDefaults docDefaults = ctStyles.addNewDocDefaults();

        CTRPrDefault ctrPrDefault = docDefaults.addNewRPrDefault();
        CTRPr ctrPr = ctrPrDefault.addNewRPr();
        if (StringUtils.isNotBlank(color)) {
            ctrPr.addNewColor().setVal(color);
        }
        if (fontSize != null) {
            ctrPr.addNewSz().setVal(new BigInteger(Integer.toString(fontSize)).multiply(new BigInteger("2")));
        }

        CTFonts ctFonts = null;
        if (StringUtils.isNotBlank(defaultFont)) {
            ctFonts = ctrPr.addNewRFonts();
            ctFonts.setAscii(defaultFont);
            ctFonts.setHAnsi(defaultFont);
            ctFonts.setCs(defaultFont);
            ctFonts.setEastAsia(eastAsiaFont);
        }

        if (StringUtils.isNotBlank(eastAsiaFont)) {
            if (ctFonts == null) {
                ctFonts = ctrPr.addNewRFonts();
            }
            ctFonts.setEastAsia(eastAsiaFont);
        }
        XWPFStyles xwpfStyles = doc.createStyles();
        xwpfStyles.setStyles(ctStyles);
        return ctStyles;
    }
}
