package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.List;

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
     * @param document {@link XWPFDocument}
     * @param rows     行数
     * @param cols     列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTableWithoutBorder(XWPFDocument document, int rows, int cols) {
        return addTable(document, rows, cols, PoiWordTool.getContentWidthOfDxa(document), XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }

    /**
     * 在单元格内创建没有边框的表格
     *
     * @param tableCell {@link XWPFTableCell}
     * @param rows      行数
     * @param cols      列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTableWithoutBorder(XWPFTableCell tableCell, int rows, int cols) {
        return addTable(tableCell, rows, cols, tableCell.getWidth(), XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }


    /**
     * 创建没有边框的表格
     *
     * @param document  {@link XWPFDocument}
     * @param rows      行数
     * @param cols      列数
     * @param isAutoFit 是否自适应宽度
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTableWithoutBorder(XWPFDocument document, int rows, int cols, boolean isAutoFit) {
        XWPFTable table = document.createTable();
        if (isAutoFit) {
            initTable(table, rows, cols, 0, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF", STTblLayoutType.AUTOFIT);
        } else {
            initTable(table, rows, cols, PoiWordTool.getContentWidthOfDxa(document), XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF", STTblLayoutType.FIXED);
        }
        return table;
    }

    /**
     * 在单元格内创建没有边框的表格
     *
     * @param tableCell {@link XWPFTableCell}
     * @param rows      行数
     * @param cols      列数
     * @param isAutoFit 是否自适应宽度
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTableWithoutBorder(XWPFTableCell tableCell, int rows, int cols, boolean isAutoFit) {
        XWPFTable table = new XWPFTable(tableCell.getCTTc().addNewTbl(), tableCell);
        tableCell.insertTable(tableCell.getTables().size(), table);
        if (isAutoFit) {
            initTable(table, rows, cols, 0, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF", STTblLayoutType.AUTOFIT);
        } else {
            initTable(table, rows, cols, tableCell.getWidth(), XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF", STTblLayoutType.FIXED);
        }
        return table;
    }

    /**
     * 创建没有边框的表格
     *
     * @param document  {@link XWPFDocument}
     * @param rows      行数
     * @param cols      列数
     * @param isAutoFit 是否自适应宽度
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFDocument document, int rows, int cols, boolean isAutoFit) {
        XWPFTable table = document.createTable();
        if (isAutoFit) {
            initTable(table, rows, cols, 0, XWPFTable.XWPFBorderType.SINGLE, 2, "000000", STTblLayoutType.AUTOFIT);
        } else {
            initTable(table, rows, cols, PoiWordTool.getContentWidthOfDxa(document), XWPFTable.XWPFBorderType.SINGLE, 2, "000000", STTblLayoutType.FIXED);
        }
        return table;
    }

    /**
     * 在单元格内创建没有边框的表格
     *
     * @param tableCell {@link XWPFTableCell}
     * @param rows      行数
     * @param cols      列数
     * @param isAutoFit 是否自适应宽度
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFTableCell tableCell, int rows, int cols, boolean isAutoFit) {
        XWPFTable table = new XWPFTable(tableCell.getCTTc().addNewTbl(), tableCell);
        tableCell.insertTable(tableCell.getTables().size(), table);
        if (isAutoFit) {
            initTable(table, rows, cols, 0, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF", STTblLayoutType.AUTOFIT);
        } else {
            initTable(table, rows, cols, tableCell.getWidth(), XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF", STTblLayoutType.FIXED);
        }
        return table;
    }

    /**
     * 创建没有边框的表格
     *
     * @param document   {@link XWPFDocument}
     * @param rows       行数
     * @param cols       列数
     * @param tableWidth 表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTableWithoutBorder(XWPFDocument document, int rows, int cols, long tableWidth) {
        return addTable(document, rows, cols, tableWidth, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }

    /**
     * 在单元格内创建没有边框的表格
     *
     * @param tableCell  {@link XWPFTableCell}
     * @param rows       行数
     * @param cols       列数
     * @param tableWidth 表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTableWithoutBorder(XWPFTableCell tableCell, int rows, int cols, long tableWidth) {
        return addTable(tableCell, rows, cols, tableWidth, XWPFTable.XWPFBorderType.NONE, 0, "FFFFFF");
    }

    /**
     * 创建表格
     *
     * @param document   {@link XWPFDocument}
     * @param rows       行数
     * @param cols       列数
     * @param tableWidth 表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFDocument document, int rows, int cols, long tableWidth) {
        return addTable(document, rows, cols, tableWidth, XWPFTable.XWPFBorderType.SINGLE, 2, "000000");
    }

    /**
     * 在单元格内创建表格
     *
     * @param tableCell  {@link XWPFTableCell}
     * @param rows       行数
     * @param cols       列数
     * @param tableWidth 表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFTableCell tableCell, int rows, int cols, long tableWidth) {
        return addTable(tableCell, rows, cols, tableWidth, XWPFTable.XWPFBorderType.SINGLE, 2, "000000");
    }

    /**
     * 创建表格
     *
     * @param document    {@link XWPFDocument}
     * @param borderSize  边框宽度
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     * @param rows        行数
     * @param cols        列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFDocument document, int rows, int cols, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        return addTable(document, rows, cols, PoiWordTool.getContentWidthOfDxa(document), borderType, borderSize, borderColor);
    }

    /**
     * 在单元格内创建表格
     *
     * @param tableCell   {@link XWPFTableCell}
     * @param borderSize  边框宽度
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     * @param rows        行数
     * @param cols        列数
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFTableCell tableCell, int rows, int cols, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        return addTable(tableCell, rows, cols, tableCell.getWidth(), borderType, borderSize, borderColor);
    }

    /**
     * 创建表格
     *
     * @param document    {@link XWPFDocument}
     * @param rows        行数
     * @param cols        列数
     * @param tableWidth  表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     * @param borderType  边框样式
     * @param borderSize  边框宽度，取值范围：[2, 96]，2：1/4 磅，96：12磅
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFDocument document, int rows, int cols, long tableWidth, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        XWPFTable table = document.createTable();
        initTable(table, rows, cols, tableWidth, borderType, borderSize, borderColor, STTblLayoutType.FIXED);
        return table;
    }

    /**
     * 在单元格内创建表格
     *
     * @param tableCell   {@link XWPFTableCell}
     * @param rows        行数
     * @param cols        列数
     * @param tableWidth  表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     * @param borderType  边框样式
     * @param borderSize  边框宽度，取值范围：[2, 96]，2：1/4 磅，96：12磅
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     *
     * @return {@link XWPFTable}
     */
    public static XWPFTable addTable(XWPFTableCell tableCell, int rows, int cols, long tableWidth, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {
        XWPFTable table = new XWPFTable(tableCell.getCTTc().addNewTbl(), tableCell);
        tableCell.insertTable(tableCell.getTables().size(), table);
        initTable(table, rows, cols, tableWidth, borderType, borderSize, borderColor, STTblLayoutType.FIXED);
        return table;
    }

    /**
     * 初始化表格
     *
     * @param table       {@link XWPFTable}
     * @param rows        行数
     * @param cols        列数
     * @param tableWidth  表格宽度，单位：DXA（可以通过 {@link PoiUnitTool#pointToDXA(double)} 将“磅”转换为 DXA）
     * @param borderType  边框样式
     * @param borderSize  边框宽度，取值范围：[2, 96]，2：1/4 磅，96：12磅
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     * @param layoutType  布局方式
     */
    public static void initTable(XWPFTable table, int rows, int cols, long tableWidth, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor, STTblLayoutType.Enum layoutType) {
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr ctTblPr;
        if ((ctTblPr = ctTbl.getTblPr()) == null) {
            ctTblPr = ctTbl.addNewTblPr();
        }
        if (layoutType != null) {
            CTTblLayoutType ctTblLayoutType;
            if ((ctTblLayoutType = ctTblPr.getTblLayout()) == null) {
                ctTblLayoutType = ctTblPr.addNewTblLayout();
            }
            ctTblLayoutType.setType(layoutType);

            if (layoutType == STTblLayoutType.FIXED) {
                table.setWidthType(TableWidthType.DXA);
                table.setWidth(String.valueOf(tableWidth));
            }
        } else {
            table.setWidthType(TableWidthType.DXA);
            table.setWidth(String.valueOf(tableWidth));
        }

        if (rows > 1) {
            for (int i = 1; i < rows; i++) {
                table.createRow();
            }
        }

        if (cols > 1) {
            XWPFTableRow tableRowOne = table.getRow(0);

            BigDecimal cellWidth = new BigDecimal(tableWidth).divide(new BigDecimal(cols), 3, RoundingMode.FLOOR);

            for (int i = 1; i < cols; i++) {
                tableRowOne.addNewTableCell();
            }

            if (layoutType == STTblLayoutType.FIXED) {
                // 固定宽度时，计算每个单元格宽度
                XWPFTableCell tableCell;
                for (int i = 0; i < cols; i++) {
                    tableCell = tableRowOne.getCell(i);
                    tableCell.setWidth(String.valueOf(cellWidth.intValue()));
                    tableCell.setWidthType(TableWidthType.DXA);
                }
            }
        }

        // 设置表格边框
        setTableBorder(table, borderType, borderSize, borderColor);
    }

    /**
     * 移除表格边框
     *
     * @param table {@link XWPFTable}
     */
    public static void removeTableBorder(XWPFTable table) {
        setTableBorder(table, XWPFTable.XWPFBorderType.NONE, 0, "000000");
    }

    /**
     * 设置表格边框
     *
     * @param table       {@link XWPFTable}
     * @param borderType  边框样式
     * @param borderSize  边框宽度，取值范围：[2, 96]，2：1/4 磅，96：12磅
     * @param borderColor 边框颜色（RGB 格式，例如："FFFFFF"）
     */
    public static void setTableBorder(XWPFTable table, XWPFTable.XWPFBorderType borderType, int borderSize, String borderColor) {

        table.setTopBorder(borderType, borderSize, 0, borderColor);
        table.setBottomBorder(borderType, borderSize, 0, borderColor);
        table.setLeftBorder(borderType, borderSize, 0, borderColor);
        table.setRightBorder(borderType, borderSize, 0, borderColor);

        table.setInsideHBorder(borderType, borderSize, 0, borderColor);
        table.setInsideVBorder(borderType, borderSize, 0, borderColor);
    }

    /**
     * 设置表格位置
     *
     * @param table      {@link XWPFTable}
     * @param horzAnchor 水平对齐锚点
     * @param leftOffset 水平偏移距离，单位：dxa
     * @param vertAnchor 垂直对齐锚点
     * @param topOffset  垂直偏移距离，单位：dxa
     */
    public static void setTablePosition(XWPFTable table, STHAnchor.Enum horzAnchor, long leftOffset,
                                        STVAnchor.Enum vertAnchor, long topOffset) {
        setTablePosition(table, horzAnchor, leftOffset, null, vertAnchor, topOffset, null, 0, 0, 0, 0);
    }


    /**
     * 设置表格位置
     *
     * @param table      {@link XWPFTable}
     * @param horzAnchor 水平对齐锚点
     * @param xAlign     水平对齐方式
     * @param vertAnchor 垂直对齐锚点
     * @param yAlign     垂直对齐方式
     */
    public static void setTablePosition(XWPFTable table, STHAnchor.Enum horzAnchor, STXAlign.Enum xAlign,
                                        STVAnchor.Enum vertAnchor, STYAlign.Enum yAlign) {
        setTablePosition(table, horzAnchor, 0, xAlign, vertAnchor, 0, yAlign, 0, 0, 0, 0);
    }


    /**
     * 设置表格位置
     *
     * @param table      {@link XWPFTable}
     * @param horzAnchor 水平对齐锚点
     * @param leftOffset 水平偏移距离，单位：dxa
     * @param xAlign     水平对齐方式
     * @param vertAnchor 垂直对齐锚点
     * @param topOffset  垂直偏移距离，单位：dxa
     * @param yAlign     垂直对齐方式
     */
    public static void setTablePosition(XWPFTable table, STHAnchor.Enum horzAnchor, long leftOffset, STXAlign.Enum xAlign,
                                        STVAnchor.Enum vertAnchor, long topOffset, STYAlign.Enum yAlign) {
        setTablePosition(table, horzAnchor, leftOffset, xAlign, vertAnchor, topOffset, yAlign, 0, 0, 0, 0);
    }

    /**
     * 设置表格位置
     *
     * @param table          {@link XWPFTable}
     * @param horzAnchor     水平对齐锚点
     * @param leftOffset     水平偏移距离，单位：dxa
     * @param xAlign         水平对齐方式
     * @param vertAnchor     垂直对齐锚点
     * @param topOffset      垂直偏移距离，单位：dxa
     * @param yAlign         垂直对齐方式
     * @param topFromText    顶部和文字的距离，单位：dxa
     * @param rightFromText  右边和文字的距离，单位：dxa
     * @param bottomFromText 底部和文字的距离，单位：dxa
     * @param leftFromText   左边和文字的距离，单位：dxa
     */
    public static void setTablePosition(XWPFTable table, STHAnchor.Enum horzAnchor, long leftOffset, STXAlign.Enum xAlign,
                                        STVAnchor.Enum vertAnchor, long topOffset, STYAlign.Enum yAlign,
                                        long topFromText, long rightFromText, long bottomFromText, long leftFromText) {
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr ctTblPr;
        if ((ctTblPr = ctTbl.getTblPr()) == null) {
            ctTblPr = ctTbl.addNewTblPr();
        }
        CTTblPPr ctTblPPr;
        if ((ctTblPPr = ctTblPr.getTblpPr()) == null) {
            ctTblPPr = ctTblPr.addNewTblpPr();
        }
        ctTblPPr.setHorzAnchor(horzAnchor);
        if (xAlign == null) {
            ctTblPPr.setTblpX(BigInteger.valueOf(leftOffset));
        } else {
            ctTblPPr.setTblpXSpec(xAlign);
        }

        ctTblPPr.setVertAnchor(vertAnchor);
        if (yAlign == null) {
            ctTblPPr.setTblpY(BigInteger.valueOf(topOffset));
        } else {
            ctTblPPr.setTblpYSpec(yAlign);
        }

        ctTblPPr.setTopFromText(BigInteger.valueOf(topFromText));
        ctTblPPr.setRightFromText(BigInteger.valueOf(rightFromText));
        ctTblPPr.setBottomFromText(BigInteger.valueOf(bottomFromText));
        ctTblPPr.setLeftFromText(BigInteger.valueOf(leftFromText));
    }

    /**
     * 设置表格行高
     *
     * @param tableRowOne {@link XWPFTableRow}
     * @param height      高度（单位：DXA）
     */
    public static void setTableRowHeight(XWPFTableRow tableRowOne, long height) {
        CTRow ctRow = tableRowOne.getCtRow();
        CTTrPr trPr = (ctRow.isSetTrPr()) ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight ctHeight = trPr.sizeOfTrHeightArray() == 0 ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
        ctHeight.setHRule(STHeightRule.EXACT);
        ctHeight.setVal(BigInteger.valueOf(height));
    }

    /**
     * 设置表格行高
     *
     * @param tableRowOne {@link XWPFTableRow}
     * @param height      高度（单位：磅）
     */
    public static void setTableRowHeightOfPoint(XWPFTableRow tableRowOne, double height) {
        CTRow ctRow = tableRowOne.getCtRow();
        CTTrPr trPr = (ctRow.isSetTrPr()) ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight ctHeight = trPr.sizeOfTrHeightArray() == 0 ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
        ctHeight.setHRule(STHeightRule.EXACT);
        ctHeight.setVal(BigInteger.valueOf(PoiUnitTool.pointToDXA(height)));
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
     * 获取单元格的第一个段落元素
     *
     * @param table   表格 {@link XWPFTable }
     * @param rowPos  行号（从 0 开始）
     * @param cellPos 列号（从 0 开始）
     */
    public static XWPFParagraph getTableCellParagraph(XWPFTable table, int rowPos, int cellPos) {
        XWPFTableCell tableCell = getTableCell(table, rowPos, cellPos);
        return getTableCellParagraph(tableCell);
    }

    /**
     * 获取单元格的第一个段落元素
     *
     * @param tableCell 单元格
     */
    public static XWPFParagraph getTableCellParagraph(XWPFTableCell tableCell) {
        List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
        if (paragraphs.size() == 0) {
            return tableCell.addParagraph();
        }
        return paragraphs.get(0);
    }

    /**
     * 获取单元格
     *
     * @param table   表格 {@link XWPFTable }
     * @param rowPos  行号（从 0 开始）
     * @param cellPos 列号（从 0 开始）
     */
    public static XWPFTableCell getTableCell(XWPFTable table, int rowPos, int cellPos) {
        XWPFTableRow tableRowOne = table.getRow(rowPos);
        while (tableRowOne == null) {
            // 创建截止到当前位置，缺少的行
            table.createRow();
            tableRowOne = table.getRow(rowPos);
        }
        XWPFTableCell cell = tableRowOne.getCell(cellPos);
        while (cell == null) {
            // 创建截止到当前位置，缺少的列
            tableRowOne.addNewTableCell();
            cell = tableRowOne.getCell(cellPos);
        }
        return cell;
    }

    /**
     * 设置单元格文字，对齐方式：左对齐；垂直居中
     *
     * @param table     表格 {@link XWPFTable }
     * @param rowPos    行号（从 0 开始）
     * @param cellPos   列号（从 0 开始）
     * @param text      文本
     * @param autoWidth 宽度自适应
     */
    public static void setTableCell(XWPFTable table, int rowPos, int cellPos, String text, boolean autoWidth) {
        XWPFTableCell tableCell = PoiWordTableTool.getTableCell(table, rowPos, cellPos);
        setTableCell(tableCell, text, autoWidth, STJc.LEFT, STVerticalJc.CENTER);
    }

    /**
     * 设置单元格文字
     *
     * @param table           表格 {@link XWPFTable }
     * @param rowPos          行号（从 0 开始）
     * @param cellPos         列号（从 0 开始）
     * @param text            文本
     * @param autoWidth       宽度自适应
     * @param horizontalAlign 水平对齐方式
     * @param verticalAlign   垂直对齐方式
     */
    public static void setTableCell(XWPFTable table, int rowPos, int cellPos, String text, boolean autoWidth, STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        XWPFTableCell tableCell = PoiWordTableTool.getTableCell(table, rowPos, cellPos);
        setTableCell(tableCell, text, autoWidth, horizontalAlign, verticalAlign);
    }

    /**
     * 设置单元格文字
     *
     * @param cell            单元格
     * @param text            单元格
     * @param autoWidth       宽度自适应
     * @param horizontalAlign 水平对齐方式
     * @param verticalAlign   垂直对齐方式
     */
    public static void setTableCell(XWPFTableCell cell, String text, boolean autoWidth, STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        cell.setText(text);
        setTableCellAlign(cell, horizontalAlign, verticalAlign);
        if (autoWidth) {
            setTableCellWidthOfAuto(cell);
        }
    }

    /**
     * 设置单元格文字
     *
     * @param cell            单元格
     * @param text            单元格
     * @param horizontalAlign 水平对齐方式
     * @param verticalAlign   垂直对齐方式
     */
    public static void setTableCellText(XWPFTableCell cell, String text, STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        cell.setText(text);
        setTableCellAlign(cell, horizontalAlign, verticalAlign);
    }

    /**
     * 设置单元格文字（文字不加粗，无下划线，单元格水平方向左对齐，垂直方向居中对齐）
     *
     * @param table           表格 {@link XWPFTable }
     * @param rowPos          行号（从 0 开始）
     * @param cellPos         列号（从 0 开始）
     * @param text            文本
     * @param fontFamily      字体
     * @param fontSize        字号（单位：磅）
     * @param color           字体颜色
     */
    public static void setTableCellText(XWPFTable table, int rowPos, int cellPos, String text, String fontFamily, Integer fontSize, String color) {
        XWPFTableCell tableCell = PoiWordTableTool.getTableCell(table, rowPos, cellPos);
        setTableCellText(tableCell, text, fontFamily, fontSize, color);
    }

    /**
     * 设置单元格文字（单元格水平方向左对齐，垂直方向居中对齐）
     *
     * @param table           表格 {@link XWPFTable }
     * @param rowPos          行号（从 0 开始）
     * @param cellPos         列号（从 0 开始）
     * @param text            文本
     * @param fontFamily      字体
     * @param fontSize        字号（单位：磅）
     * @param color           字体颜色
     * @param bold            是否加粗
     * @param underline       是否加下划线
     */
    public static void setTableCellText(XWPFTable table, int rowPos, int cellPos, String text, String fontFamily, Integer fontSize, String color,
                                        boolean bold, boolean underline) {
        XWPFTableCell tableCell = PoiWordTableTool.getTableCell(table, rowPos, cellPos);
        setTableCellText(tableCell, text, fontFamily, fontSize, color, bold, underline);
    }

    /**
     * 设置单元格文字
     *
     * @param table           表格 {@link XWPFTable }
     * @param rowPos          行号（从 0 开始）
     * @param cellPos         列号（从 0 开始）
     * @param text            文本
     * @param fontFamily      字体
     * @param fontSize        字号（单位：磅）
     * @param color           字体颜色
     * @param bold            是否加粗
     * @param underline       是否加下划线
     * @param horizontalAlign 单元格水平对齐方式
     * @param verticalAlign   单元格垂直对齐方式
     */
    public static void setTableCellText(XWPFTable table, int rowPos, int cellPos, String text, String fontFamily, Integer fontSize, String color,
                                        boolean bold, boolean underline,
                                        STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        XWPFTableCell tableCell = PoiWordTableTool.getTableCell(table, rowPos, cellPos);
        setTableCellText(tableCell, text, fontFamily, fontSize, color, bold, underline, horizontalAlign, verticalAlign);
    }

    /**
     * 设置单元格文字（文字不加粗，无下划线，单元格水平方向左对齐，垂直方向居中对齐）
     *
     * @param tableCell       单元格
     * @param text            文本
     * @param fontFamily      字体
     * @param fontSize        字号（单位：磅）
     * @param color           字体颜色
     */
    public static void setTableCellText(XWPFTableCell tableCell, String text, String fontFamily, Integer fontSize, String color) {
        setTableCellText(tableCell, text, fontFamily, fontSize, color, false, false, STJc.LEFT, STVerticalJc.CENTER);
    }

    /**
     * 设置单元格文字（单元格水平方向左对齐，垂直方向居中对齐）
     *
     * @param tableCell       单元格
     * @param text            文本
     * @param fontFamily      字体
     * @param fontSize        字号（单位：磅）
     * @param color           字体颜色
     * @param bold            是否加粗
     * @param underline       是否加下划线
     */
    public static void setTableCellText(XWPFTableCell tableCell, String text, String fontFamily, Integer fontSize, String color,
                                        boolean bold, boolean underline) {
        setTableCellText(tableCell, text, fontFamily, fontSize, color, bold, underline, STJc.LEFT, STVerticalJc.CENTER);
    }

    /**
     * 设置单元格文字
     *
     * @param tableCell       单元格
     * @param text            文本
     * @param fontFamily      字体
     * @param fontSize        字号（单位：磅）
     * @param color           字体颜色
     * @param bold            是否加粗
     * @param underline       是否加下划线
     * @param horizontalAlign 单元格水平对齐方式
     * @param verticalAlign   单元格垂直对齐方式
     */
    public static void setTableCellText(XWPFTableCell tableCell, String text,
                                        String fontFamily, Integer fontSize, String color,
                                        boolean bold, boolean underline,
                                        STJc.Enum horizontalAlign, STVerticalJc.Enum verticalAlign) {
        XWPFParagraph paragraph = getTableCellParagraph(tableCell);
        PoiWordParagraphTool.addTxt(paragraph, text, fontFamily, fontSize, color, bold, underline);
        setTableCellAlign(tableCell, horizontalAlign, verticalAlign);
    }

    /**
     * 设置单元格宽度
     *
     * @param tableCell 单元格
     * @param width     宽度（单位：DXA）
     */
    public static void setTableCellWidth(XWPFTableCell tableCell, long width) {
        tableCell.setWidth(String.valueOf(width));
        tableCell.setWidthType(TableWidthType.DXA);
    }

    /**
     * 设置单元格宽度自适应
     *
     * @param tableCell 单元格
     */
    public static void setTableCellWidthOfAuto(XWPFTableCell tableCell) {
        tableCell.setWidthType(TableWidthType.AUTO);
    }

    /**
     * 设置单元格宽度
     *
     * @param tableCell 单元格
     * @param width     宽度（单位：磅）
     */
    public static void setTableCellWidthOfPoint(XWPFTableCell tableCell, double width) {
        tableCell.setWidth(String.valueOf(PoiUnitTool.pointToDXA(width)));
        tableCell.setWidthType(TableWidthType.DXA);
    }

    /**
     * 设置单元格对齐方式
     *
     * @param cell            单元格
     * @param horizontalAlign 水平对齐方式
     * @param verticalAlign   垂直对齐方式
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

    private static CTTcBorders getCTTcBorders(XWPFTableCell tableCell) {
        CTTc ctTc = tableCell.getCTTc();
        CTTcPr ctTcPr;
        if ((ctTcPr = ctTc.getTcPr()) == null) {
            ctTcPr = ctTc.addNewTcPr();
        }
        CTTcBorders ctTblBorders;
        if ((ctTblBorders = ctTcPr.getTcBorders()) == null) {
            ctTblBorders = ctTcPr.addNewTcBorders();
        }
        return ctTblBorders;
    }

    /**
     * 设置单元格上边框宽度
     *
     * @param tableCell 单元格
     * @param width     宽度，单位：1/8磅，取值：[2, 96]，即：1/4磅 到 12磅
     * @param color     颜色
     * @param style     样式
     */
    public static void setTableCellBorderOfTop(XWPFTableCell tableCell, int width, String color, STBorder.Enum style) {
        CTTcBorders ctTblBorders = getCTTcBorders(tableCell);
        CTBorder ctBorder;
        if ((ctBorder = ctTblBorders.getTop()) == null) {
            ctBorder = ctTblBorders.addNewTop();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
    }

    /**
     * 设置单元格下边框宽度
     *
     * @param tableCell 单元格
     * @param width     宽度，单位：1/8磅，取值：[2, 96]，即：1/4磅 到 12磅
     * @param color     颜色
     * @param style     样式
     */
    public static void setTableCellBorderOfBottom(XWPFTableCell tableCell, int width, String color, STBorder.Enum style) {
        CTTcBorders ctTblBorders = getCTTcBorders(tableCell);
        CTBorder ctBorder;
        if ((ctBorder = ctTblBorders.getBottom()) == null) {
            ctBorder = ctTblBorders.addNewBottom();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
    }

    /**
     * 设置单元格左边框宽度
     *
     * @param tableCell 单元格
     * @param width     宽度，单位：1/8磅，取值：[2, 96]，即：1/4磅 到 12磅
     * @param color     颜色
     * @param style     样式
     */
    public static void setTableCellBorderOfLeft(XWPFTableCell tableCell, int width, String color, STBorder.Enum style) {
        CTTcBorders ctTblBorders = getCTTcBorders(tableCell);
        CTBorder ctBorder;
        if ((ctBorder = ctTblBorders.getLeft()) == null) {
            ctBorder = ctTblBorders.addNewLeft();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
    }

    /**
     * 设置单元格右边框宽度
     *
     * @param tableCell 单元格
     * @param width     宽度，单位：1/8磅，取值：[2, 96]，即：1/4磅 到 12磅
     * @param color     颜色
     * @param style     样式
     */
    public static void setTableCellBorderOfRight(XWPFTableCell tableCell, int width, String color, STBorder.Enum style) {
        CTTcBorders ctTblBorders = getCTTcBorders(tableCell);
        CTBorder ctBorder;
        if ((ctBorder = ctTblBorders.getRight()) == null) {
            ctBorder = ctTblBorders.addNewRight();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
    }

    /**
     * 设置单元格宽度
     *
     * @param tableCell 单元格
     * @param width     宽度，单位：1/8磅，取值：[2, 96]，即：1/4磅 到 12磅
     * @param color     颜色
     * @param style     样式
     */
    public static void setTableCellBorder(XWPFTableCell tableCell, int width, String color, STBorder.Enum style) {
        CTTcBorders ctTblBorders = getCTTcBorders(tableCell);
        CTBorder ctBorder;
        if ((ctBorder = ctTblBorders.getTop()) == null) {
            ctBorder = ctTblBorders.addNewTop();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
        if ((ctBorder = ctTblBorders.getBottom()) == null) {
            ctBorder = ctTblBorders.addNewBottom();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
        if ((ctBorder = ctTblBorders.getLeft()) == null) {
            ctBorder = ctTblBorders.addNewLeft();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
        if ((ctBorder = ctTblBorders.getRight()) == null) {
            ctBorder = ctTblBorders.addNewRight();
        }
        ctBorder.setSz(BigInteger.valueOf(width));
        ctBorder.setColor(color);
        ctBorder.setVal(style);
    }

    /**
     * 设置单元格背景色
     *
     * @param cell    单元格
     * @param bgColor 背景色（RGB 格式，例如："FFFFFF"）
     */
    public static void setTableCellBgColor(XWPFTableCell cell, String bgColor) {
        cell.setColor(bgColor);
    }
}
