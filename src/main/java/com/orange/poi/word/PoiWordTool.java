package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.officeDocument.x2006.extendedProperties.CTProperties;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STDocGrid;

import java.io.File;
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
     * 默认字体大小，单位：磅
     */
    public static final int DEFAULT_FONT_SIZE = 12;
    /**
     * 基准行高
     */
    public static final int LINE_HEIGHT_DXA   = 240;

    /**
     * A4 纸宽度，单位：DXA
     */
    public static final int    A4_WIDTH_DXA            = 11906;
    /**
     * A4 纸高度，单位：DXA
     */
    public static final int    A4_HEIGHT_DXA           = 16838;
    /**
     * 页面上边距默认为 1.5 cm
     */
    public static final long   A4_MARGIN_TOP_DXA       = PoiUnitTool.centimeterToDXA(1.5f);
    /**
     * 页面右边距默认为 1.5 cm
     */
    public static final long   A4_MARGIN_RIGHT_DXA     = PoiUnitTool.centimeterToDXA(1.5f);
    /**
     * 页面下边距默认为 1.5 cm
     */
    public static final long   A4_MARGIN_BOTTOM_DXA    = PoiUnitTool.centimeterToDXA(1.5f);
    /**
     * 页面左边距默认为 1.5 cm
     */
    public static final long   A4_MARGIN_LEFT_DXA      = PoiUnitTool.centimeterToDXA(1.5f);
    /**
     * 内容宽度（页面宽度减去左右页边距），单位：DXA
     */
    public static final long   A4_CONTENT_WIDTH_DXA    = A4_WIDTH_DXA - A4_MARGIN_LEFT_DXA - A4_MARGIN_RIGHT_DXA;
    /**
     * 内容宽度（页面宽度减去左右页边距），单位：磅
     */
    public static final double A4_CONTENT_WIDTH_POINT  = PoiUnitTool.dxaToPoint(A4_CONTENT_WIDTH_DXA);
    /**
     * 内容高度（页面高度减去上下页边距），单位：DXA
     */
    public static final long   A4_CONTENT_HEIGHT_DXA   = A4_HEIGHT_DXA - A4_MARGIN_TOP_DXA - A4_MARGIN_BOTTOM_DXA;
    /**
     * 内容高度（页面高度减去上下页边距），单位：磅
     */
    public static final double A4_CONTENT_HEIGHT_POINT = PoiUnitTool.dxaToPoint(A4_CONTENT_HEIGHT_DXA);
    /**
     * 临时文件目录
     */
    public static final File   TEMP_FILE_DIR           = new File(System.getProperty("java.io.tmpdir"));

    /**
     * 将 word 初始化为 A4 值（包括页面大小、页边距等）
     */
    public static void initDocForA4(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageSz pageSize = sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(A4_WIDTH_DXA));
        pageSize.setH(BigInteger.valueOf(A4_HEIGHT_DXA));

        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setTop(BigInteger.valueOf(A4_MARGIN_TOP_DXA));
        pageMar.setRight(BigInteger.valueOf(A4_MARGIN_RIGHT_DXA));
        pageMar.setBottom(BigInteger.valueOf(A4_MARGIN_BOTTOM_DXA));
        pageMar.setLeft(BigInteger.valueOf(A4_MARGIN_LEFT_DXA));

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
