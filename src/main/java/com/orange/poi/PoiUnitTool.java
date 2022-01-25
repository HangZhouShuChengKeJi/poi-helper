package com.orange.poi;

import org.apache.poi.util.Units;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

/**
 * apache poi 单位换算 工具类
 *
 * @author 小天
 * @date 2019/5/28 1:45
 */
public class PoiUnitTool {

    /**
     * 厘米 转换为 磅（像素点数）
     *
     * @param cm 厘米
     *
     * @return 磅（像素点数）
     */
    public static double centimeterToPoint(double cm) {
        // 1 英寸 = 2.54 厘米
        // 网络图片的分辨率一般是 72dpi，即：1 英寸有 72 个像素点
        return (cm / 2.54f) * Units.POINT_DPI;
    }

    /**
     * 厘米 转换为 {@link STTblWidth#DXA}
     *
     * @param cm 厘米
     *
     * @return {@link STTblWidth#DXA} 厘米
     */
    public static long centimeterToDXA(double cm) {
        return (long) (centimeterToPoint(cm) * 20);
    }

    /**
     * 厘米 转换为 像素
     *
     * @param cm 厘米
     *
     * @return 像素
     */
    public static long centimeterToPixel(double cm) {
        return pointToPixel(centimeterToPoint(cm));
    }

    /**
     * 磅（像素点数） 转换为 像素
     *
     * @param point 像素点数
     *
     * @return 像素
     */
    public static int pointToPixel(double point) {
        return Units.pointsToPixel(point);
    }

    /**
     * 磅（像素点数） 转换为 {@link STTblWidth#DXA}
     *
     * @param point 磅（像素点数）
     *
     * @return {@link STTblWidth#DXA} 值
     */
    public static long pointToDXA(double point) {
        return (long) (point * 20);
    }


    /**
     * {@link STTblWidth#DXA} 转换为 磅（像素点数）
     *
     * @param dxa {@link STTblWidth#DXA}
     *
     * @return 磅（像素点数）
     */
    public static double dxaToPoint(long dxa) {
        return dxa / 20.00d;
    }

    /**
     * {@link STTblWidth#DXA} 转换为 像素
     *
     * @param dxa {@link STTblWidth#DXA}
     *
     * @return 像素
     */
    public static int dxaToPixel(long dxa) {
        return Units.pointsToPixel(dxaToPoint(dxa));
    }

    /**
     * 像素 转换为 {@link STTblWidth#DXA}
     *
     * @param pixel 像素
     *
     * @return {@link STTblWidth#DXA} 值
     */
    public static long pixelToDXA(int pixel) {
        double points = Units.pixelToPoints(pixel);
        return (long) (points * 20);
    }

    /**
     * 像素 转换为 磅（像素点数）
     *
     * @param pixel 像素
     *
     * @return 磅（像素点数）
     */
    public static double pixelToPoint(int pixel) {
        return Units.pixelToPoints(pixel);
    }

    /**
     * 像素 转换为 emu
     *
     * @param pixel 像素
     *
     * @return emu
     */
    public static int pixelToEMU(int pixel) {
        return Units.toEMU(pixelToPoint(pixel));
    }

    /**
     * emu 转换为 像素
     *
     * @param emu emu
     *
     * @return 像素
     */
    public static int emuToPixel(long emu) {
        return pointToPixel(Units.toPoints(emu));
    }

}
