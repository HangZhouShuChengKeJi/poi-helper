package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.paper.PaperSize;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STOnOff1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;

import java.math.BigInteger;

/**
 * section 元素工具。
 *
 * @author 小天
 * @date 2022/1/25 16:05
 */
public class PoiWordSectionTool {

    /**
     * 获取 word 文档 默认的节属性
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @return 节对象 {@link CTSectPr}
     */
    public static CTSectPr getSectPr(XWPFDocument doc) {
        CTBody body = doc.getDocument().getBody();
        if (body.isSetSectPr()) {
            return body.getSectPr();
        } else {
            return body.addNewSectPr();
        }
    }

    /**
     * 添加 section。 section 在形式上是一个 "设置了 'sectPr'" 的段落（{@link XWPFParagraph}）。
     * <p>
     * 该方法会创建一个新的段落，并设置 'sectPr'。'sectPr' 的 pgSz, pgMar, docGrid 属性与默认的 'sectPr' 保持一致。
     * <p>
     * 如有其它需要，可以通过 {@link #getSectionProperties} 方法获取 section 属性，并进行相关设置。
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @return 段落 {@link XWPFParagraph}
     */
    public static XWPFParagraph addSection(XWPFDocument doc) {
        XWPFParagraph paragraph = doc.createParagraph();
        CTSectPr docSectPr = doc.getDocument().getBody().getSectPr();
        CTSectPr ctSectPr = CTSectPr.Factory.newInstance();

        ctSectPr.setPgSz((CTPageSz)docSectPr.getPgSz().copy());
        ctSectPr.setPgMar((CTPageMar) docSectPr.getPgMar().copy());
        ctSectPr.setDocGrid((CTDocGrid) docSectPr.getDocGrid().copy());

        CTPPr ctpPr = PoiWordParagraphTool.getParagraphProperties(paragraph);
        ctpPr.setSectPr(ctSectPr);
        return paragraph;
    }

    /**
     * 获取 section 属性。
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @return {@link CTSectPr}
     */
    public static CTSectPr getSectionProperties(XWPFDocument doc) {
        return getSectionProperties(doc, true);
    }

    /**
     * 获取 section 属性。
     *
     * @param doc    文档 {@link XWPFDocument}
     * @param create 未设置 section 时，是否创建
     *
     * @return {@link CTSectPr}
     */
    public static CTSectPr getSectionProperties(XWPFDocument doc, boolean create) {
        CTBody ctBody = doc.getDocument().getBody();
        if(ctBody.isSetSectPr()) {
            return ctBody.getSectPr();
        }
        if(create) {
            return ctBody.addNewSectPr();
        }
        return null;
    }

    /**
     * 获取 section 属性。
     *
     * @param paragraph 段落 {@link XWPFParagraph}
     *
     * @return {@link CTSectPr}
     */
    public static CTSectPr getSectionProperties(XWPFParagraph paragraph) {
        CTPPr ctpPr = PoiWordParagraphTool.getParagraphProperties(paragraph);
        if (ctpPr.isSetSectPr()) {
            return ctpPr.getSectPr();
        }
        return null;
    }

    /**
     * 取消分栏
     *
     * @param ctSectPr
     */
    public static void unsetCols(CTSectPr ctSectPr) {
        if (ctSectPr.isSetCols()) {
            ctSectPr.unsetCols();
        }
    }

    /**
     * 设置等宽分栏
     *
     * @param doc       文档 {@link XWPFDocument}
     * @param colSize   分栏数量
     * @param space     分栏间距。单位：磅
     * @param splitLine 是否显示分割线
     */
    public static void setCols(XWPFDocument doc, int colSize, int space, boolean splitLine) {
        CTSectPr ctSectPr = getSectionProperties(doc);
        setCols(ctSectPr, colSize, space, splitLine);
    }

    /**
     * 设置等宽分栏
     *
     * @param ctSectPr  section 属性 {@link CTSectPr}
     * @param colSize   分栏数量
     * @param space     分栏间距。单位：磅
     * @param splitLine 是否显示分割线
     */
    public static void setCols(CTSectPr ctSectPr, int colSize, int space, boolean splitLine) {
        if (ctSectPr.isSetCols()) {
            ctSectPr.unsetCols();
        }
        CTColumns ctColumns = ctSectPr.addNewCols();
        ctColumns.setNum(BigInteger.valueOf(colSize));
        if (colSize == 1) {
            return;
        }
        ctColumns.setSpace(PoiUnitTool.pointToDXA(space));
        ctColumns.setEqualWidth(STOnOff1.ON);
        if (splitLine) {
            ctColumns.setSep(STOnOff1.ON);
        }
    }

    /**
     * 设置 section 类型。
     *
     * type 说明：
     * <ul>
     *     <li>{@link STSectionMark#NEXT_PAGE} : 下一页分节符（如果未指定类型，则为默认值），从新的页开始新的节。</li>
     *     <li>{@link STSectionMark#ODD_PAGE} : 奇数页分节符，在下一个奇数页上开始新的节。</li>
     *     <li>{@link STSectionMark#EVEN_PAGE} : 偶数页分节符，从下一个偶数页开始新的节。</li>
     *     <li>{@link STSectionMark#CONTINUOUS} : 连续的分节符，从新的段落开始新的节。</li>
     *     <li>{@link STSectionMark#NEXT_COLUMN} : 连续的分节符，从下一栏开始新的节（用于多栏场景）。</li>
     * </ul>
     *
     * @param ctSectPr section 属性 {@link CTSectPr}
     * @param type     类型 {@link CTSectType}
     */
    public static void setType(CTSectPr ctSectPr, STSectionMark.Enum type) {
        ctSectPr.addNewType().setVal(type);
    }

    /**
     * 获取页边距对象
     *
     * @param doc 文档 {@link XWPFDocument}
     *
     * @return 页边距对象
     */
    @Deprecated
    public static CTPageMar getPageMar(XWPFDocument doc) {
        CTSectPr ctSectPr = getSectPr(doc);
        return getPageMar(ctSectPr);
    }

    /**
     * 获取页边距对象
     *
     * @param ctSectPr section 属性 {@link CTSectPr}
     *
     * @return 页边距对象
     */
    public static CTPageMar getPageMar(CTSectPr ctSectPr) {
        if (ctSectPr.isSetPgMar()) {
            return ctSectPr.getPgMar();
        } else {
            return ctSectPr.addNewPgMar();
        }
    }

    /**
     * 设置页边距
     *
     * @param doc             文档 ({@link XWPFDocument})
     * @param marginTopDxa    上边距（单位：dxa）
     * @param marginBottomDxa 下边距（单位：dxa）
     * @param marginLeftDxa   左边距（单位：dxa）
     * @param marginRightDxa  右边距（单位：dxa）
     */
    @Deprecated
    public static void setPageMargin(XWPFDocument doc,
                                     long marginTopDxa, long marginBottomDxa,
                                     long marginLeftDxa, long marginRightDxa) {
        CTPageMar pageMar = getPageMar(doc);
        setPageMargin(pageMar, marginTopDxa, marginBottomDxa, marginLeftDxa, marginRightDxa);
    }

    /**
     * 设置页边距
     *
     * @param pageMar         页边距对象 ({@link CTPageMar})
     * @param marginTopDxa    上边距（单位：dxa）
     * @param marginBottomDxa 下边距（单位：dxa）
     * @param marginLeftDxa   左边距（单位：dxa）
     * @param marginRightDxa  右边距（单位：dxa）
     */
    public static void setPageMargin(CTPageMar pageMar,
                                     long marginTopDxa, long marginBottomDxa,
                                     long marginLeftDxa, long marginRightDxa) {
        pageMar.setTop(BigInteger.valueOf(marginTopDxa));
        pageMar.setBottom(BigInteger.valueOf(marginBottomDxa));
        pageMar.setLeft(BigInteger.valueOf(marginLeftDxa));
        pageMar.setRight(BigInteger.valueOf(marginRightDxa));
    }

    /**
     * 设置页边距
     *
     * @param ctSectPr        节 {@link CTSectPr}
     * @param marginTopDxa    上边距（单位：dxa）
     * @param marginBottomDxa 下边距（单位：dxa）
     * @param marginLeftDxa   左边距（单位：dxa）
     * @param marginRightDxa  右边距（单位：dxa）
     */
    public static void setPageMargin(CTSectPr ctSectPr,
                                     long marginTopDxa, long marginBottomDxa,
                                     long marginLeftDxa, long marginRightDxa) {
        CTPageMar ctPageMar = getPageMar(ctSectPr);
        setPageMargin(ctPageMar, marginTopDxa, marginBottomDxa, marginLeftDxa, marginRightDxa);
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param doc word 文档 {@link XWPFDocument}
     *
     * @return 页面宽度
     */
    @Deprecated
    public static CTPageSz getPageSize(XWPFDocument doc) {
        CTSectPr sectPr = getSectPr(doc);
        return getPageSize(sectPr);
    }

    /**
     * 获取页面宽度，单位：dxa
     *
     * @param ctSectPr section 属性 {@link CTSectPr}
     *
     * @return 页面宽度
     */
    public static CTPageSz getPageSize(CTSectPr ctSectPr) {
        CTPageSz pageSize = ctSectPr.getPgSz();
        if (ctSectPr.isSetPgSz()) {
            return pageSize;
        }
        return ctSectPr.addNewPgSz();
    }

    /**
     * 设置指定节的页面尺寸
     *
     * @param ctSectPr     节 {@link CTSectPr}
     * @param paperSize    纸张尺寸
     * @param marginTop    上边距（单位：毫米）
     * @param marginBottom 下边距（单位：毫米）
     * @param marginLeft   左边距（单位：毫米）
     * @param marginRight  右边距（单位：毫米）
     */
    public static void setPageSize(CTSectPr ctSectPr, PaperSize paperSize,
                                   float marginTop, float marginBottom, float marginLeft, float marginRight) {
        CTPageSz pageSize = getPageSize(ctSectPr);
        pageSize.setW(BigInteger.valueOf(paperSize.width_dxa));
        pageSize.setH(BigInteger.valueOf(paperSize.height_dxa));

        final long marginTopDxa = PoiUnitTool.centimeterToDXA(marginTop / 10.f);
        final long marginBottomDxa = PoiUnitTool.centimeterToDXA(marginBottom / 10.f);
        final long marginLeftDxa = PoiUnitTool.centimeterToDXA(marginLeft / 10.f);
        final long marginRightDxa = PoiUnitTool.centimeterToDXA(marginRight / 10.f);

        CTPageMar ctPageMar = getPageMar(ctSectPr);
        setPageMargin(ctPageMar, marginTopDxa, marginBottomDxa, marginLeftDxa, marginRightDxa);
    }
}
