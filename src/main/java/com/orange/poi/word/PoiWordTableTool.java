package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.List;

import static com.orange.poi.word.PoiWordTool.A4_CONTENT_WIDTH_DXA;

/**
 * apache poi word 表格工具类
 *
 * @author 小天
 * @date 2019/6/3 19:34
 */
public class PoiWordTableTool {

    /**
     * 创建没有边框的表格
     *
     * @param document {@link XWPFParagraph}
     * @param rows     行数
     * @param cols     列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTableWithoutBorder(XWPFDocument document, int rows, int cols) {
        return createTable(document, rows, cols, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }

    /**
     * 创建没有边框的表格
     *
     * @param document   {@link XWPFParagraph}
     * @param rows       行数
     * @param cols       列数
     * @param tableWidth 表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTableWithoutBorder(XWPFDocument document, int rows, int cols, int tableWidth) {
        return createTable(document, rows, cols, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }


    /**
     * 创建表格
     *
     * @param document   {@link XWPFParagraph}
     * @param rows       行数
     * @param cols       列数
     * @param tableWidth 表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTable(XWPFDocument document, int rows, int cols, int tableWidth) {
        return createTable(document, rows, cols, tableWidth, XWPFTable.XWPFBorderType.SINGLE, 2, "000000");
    }

    /**
     * 创建表格
     *
     * @param document    {@link XWPFParagraph}
     * @param borderSize  边框宽度
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     * @param rows        行数
     * @param cols        列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTable(XWPFDocument document, int rows, int cols, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        return createTable(document, rows, cols, A4_CONTENT_WIDTH_DXA, borderType, borderSize, borderColor);
    }

    /**
     * 创建表格
     *
     * @param document    {@link XWPFParagraph}
     * @param rows        行数
     * @param cols        列数
     * @param tableWidth  表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     * @param borderType  边框样式
     * @param borderSize  边框宽度，取值范围：[2, 96]，2：1/4 磅，96：12磅
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable createTable(XWPFDocument document, int rows, int cols, int tableWidth, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        XWPFTable table = document.createTable();
        table.setWidthType(TableWidthType.DXA);
        table.setWidth(String.valueOf(tableWidth));

        table.setTopBorder(borderType, borderSize, 0, borderColor);
        table.setBottomBorder(borderType, borderSize, 0, borderColor);
        table.setLeftBorder(borderType, borderSize, 0, borderColor);
        table.setRightBorder(borderType, borderSize, 0, borderColor);

        table.setInsideHBorder(borderType, borderSize, 0, borderColor);
        table.setInsideVBorder(borderType, borderSize, 0, borderColor);


        if (rows > 1) {
            for (int i = 1; i < rows; i++) {
                table.createRow();
            }
        }

        if (cols > 1) {
            XWPFTableRow tableRowOne = table.getRow(0);

            BigDecimal cellWidth = new BigDecimal(tableWidth).divide(new BigDecimal(cols), 3, RoundingMode.FLOOR);

            XWPFTableCell tableCell = tableRowOne.getCell(0);
            tableCell.setWidth(String.valueOf(cellWidth.intValue()));
            tableCell.setWidthType(TableWidthType.DXA);

            for (int i = 1; i < cols; i++) {
                tableCell = tableRowOne.addNewTableCell();
                tableCell.setWidth(String.valueOf(cellWidth.intValue()));
                tableCell.setWidthType(TableWidthType.DXA);
            }
        }
        return table;
    }

    /**
     * 设置表格行高
     *
     * @param tableRowOne {@link XWPFTableRow}
     * @param pixel       高度（单位：像素）
     */
    public static void setTableRowHeightOfPixel(XWPFTableRow tableRowOne, int pixel) {
        CTRow ctRow = tableRowOne.getCtRow();
        CTTrPr trPr = (ctRow.isSetTrPr()) ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight ctHeight = trPr.sizeOfTrHeightArray() == 0 ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
        ctHeight.setHRule(STHeightRule.EXACT);
        ctHeight.setVal(BigInteger.valueOf(PoiUnitTool.pixelToDXA(pixel)));
    }

    /**
     * 设置单元格宽度
     *
     * @param tableCell 单元格
     * @param width     宽度（单位：磅）
     */
    public static void setTableCellWidth(XWPFTableCell tableCell, int width) {
        tableCell.setWidth(String.valueOf(PoiUnitTool.pointToDXA(width)));
        tableCell.setWidthType(TableWidthType.DXA);
    }

    /**
     * 设置单元格文字
     *
     * @param cell            单元格
     * @param horizontalAlign 水平对齐方式
     * @param verticalAlign   垂直对齐方式
     */
    public static void setTableCellAlign(XWPFTableCell cell, String text, String color, STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        cell.setText(text);
        cell.setColor(color);
        setTableCellAlign(cell, horizontalAlign, verticalAlign);
    }

    /**
     * 设置单元格对齐方式
     *
     * @param cell            单元格
     * @param horizontalAlign 垂直对齐方式
     * @param verticalAlign   水平对齐方式
     */
    public static void setTableCellAlign(XWPFTableCell cell, STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        CTTc ctTc = cell.getCTTc();
        CTPPr ctpPr;

        // 垂直对齐
        if (verticalAlign != null) {
            CTTcPr tcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
            CTVerticalJc vJc = tcPr.isSetVAlign() ? tcPr.getVAlign() : tcPr.addNewVAlign();
            vJc.setVal(verticalAlign);
        }

        // 水平对齐
        if (horizontalAlign != null) {
            List<CTP> ctpList = ctTc.getPList();
            if (ctpList == null || ctpList.size() == 0) {
                return;
            }
            if ((ctpPr = ctpList.get(0).getPPr()) == null) {
                ctpPr = ctpList.get(0).addNewPPr();
            }
            if (ctpPr.isSetJc()) {
                ctpPr.getJc().setVal(horizontalAlign);
            } else {
                ctpPr.addNewJc().setVal(horizontalAlign);
            }
        }

    }

    /**
     * 设置单元格背景色
     *
     * @param cell    单元格
     * @param bgColor 背景色（RGB 格式，例如："FFFFFF"）
     */
    public static void setTableCellBgColor(XWPFTableCell cell, String bgColor) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr cellProperties;
        if ((cellProperties = ctTc.getTcPr()) == null) {
            cellProperties = ctTc.addNewTcPr();
        }
        CTShd ctShd;
        if ((ctShd = cellProperties.getShd()) == null) {
            ctShd = cellProperties.addNewShd();
        }
        ctShd.setColor(bgColor);
        ctShd.setFill(bgColor);
    }
}
