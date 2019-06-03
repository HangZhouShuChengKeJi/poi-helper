package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

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
    public static final  int DEFAULT_FONT_SIZE = 12;
    /**
     * 常量，在文档中定义长度和高度的单位
     */
    private static final int PER_LINE          = 100;
    /**
     * 基准行高
     */
    public static final  int LINE_HEIGHT_DXA   = 240;

    /**
     * A4 纸宽度，单位：DXA
     */
    public static final int  A4_WIDTH_DXA          = 11906;
    /**
     * A4 纸高度，单位：DXA
     */
    public static final int  A4_HEIGHT_DXA         = 16838;
    /**
     * 页面上边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_TOP_DXA     = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 页面右边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_RIGHT_DXA   = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 页面下边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_BOTTOM_DXA  = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 页面左边距默认为 1.5 cm
     */
    public static final int  A4_MARGIN_LEFT_DXA    = PoiUnitTool.centimeterToDXA(1.5f).intValue();
    /**
     * 内容宽度（页面宽度减去左右页边距），单位：DXA
     */
    public static final int  A4_CONTENT_WIDTH_DXA  = A4_WIDTH_DXA - A4_MARGIN_LEFT_DXA - A4_MARGIN_RIGHT_DXA;
    /**
     * 内容高度（页面高度减去上下页边距），单位：DXA
     */
    public static final int  A4_CONTENT_HEIGHT_DXA = A4_HEIGHT_DXA - A4_MARGIN_TOP_DXA - A4_MARGIN_BOTTOM_DXA;
    /**
     * 临时文件目录
     */
    public static final File TEMP_FILE_DIR         = new File(System.getProperty("java.io.tmpdir"));

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
    }



}
