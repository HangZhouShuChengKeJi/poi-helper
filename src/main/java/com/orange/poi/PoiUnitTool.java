package com.orange.poi;

import org.apache.poi.util.Units;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.math.BigInteger;

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
        return (cm / 2.54) * Units.POINT_DPI;
    }

    /**
     * 厘米 转换为 {@link STTblWidth#DXA}
     *
     * @param cm 厘米
     *
     * @return {@link STTblWidth#DXA} 厘米
     */
    public static BigInteger centimeterToDXA(double cm) {
        return BigInteger.valueOf((long) (centimeterToPoint(cm) * 20));
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
    public static BigInteger pointToDXA(double point) {
        return BigInteger.valueOf((long) (point * 20));
    }


    /**
     * {@link STTblWidth#DXA} 转换为 磅（像素点数）
     *
     * @param dxa {@link STTblWidth#DXA}
     *
     * @return 磅（像素点数）
     */
    public static int dxaToPoint(int dxa) {
        return dxa / 20;
    }

    /**
     * 像素 转换为 {@link STTblWidth#DXA}
     *
     * @param pixel 像素
     *
     * @return {@link STTblWidth#DXA} 值
     */
    public static int pixelToDXA(int pixel) {
        double points = Units.pixelToPoints(pixel);
        return (int) (points * 20);
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

}
