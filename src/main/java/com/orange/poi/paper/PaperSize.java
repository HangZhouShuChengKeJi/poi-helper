package com.orange.poi.paper;

import com.orange.poi.PoiUnitTool;

/**
 * 纸张尺寸枚举
 *
 * @author 小天
 * @date 2019/8/16 11:53
 */
public enum PaperSize {
    /**
     * A4 纸，宽：210mm，高：297mm
     */
    A4(210, 297),
    /**
     * B5 纸，宽：182mm，高：257mm
     */
    B5(182, 257);

    /**
     * 页面宽度。单位：毫米
     */
    public int  width;
    /**
     * 页面高度。单位：毫米
     */
    public int  height;
    /**
     * 页面宽度。单位：厘米
     */
    public double width_cm;
    /**
     * 页面高度。单位：厘米
     */
    public double height_cm;
    /**
     * 页面宽度。单位：dxa
     */
    public long width_dxa;
    /**
     * 页面高度。单位：dxa
     */
    public long height_dxa;
    /**
     * 页面宽度。单位：point
     */
    public double width_point;
    /**
     * 页面高度。单位：point
     */
    public double height_point;

    /**
     * 构造方法
     *
     * @param width  页面宽度，单位：毫米
     * @param height 页面高度，单位：毫米
     */
    PaperSize(int width, int height) {
        this.width = width;
        this.height = height;

        this.width_cm = width / 10.00f;
        this.height_cm = height / 10.00f;

        this.width_dxa = PoiUnitTool.centimeterToDXA(width_cm);
        this.height_dxa = PoiUnitTool.centimeterToDXA(height_cm);

        this.width_point = PoiUnitTool.centimeterToPoint(width_cm);
        this.height_point = PoiUnitTool.centimeterToPoint(height_cm);
    }
}
