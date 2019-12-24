package com.orange.poi.util;

import com.orange.poi.PoiUnitTool;
import com.sun.imageio.plugins.png.PNGImageReader;
import com.sun.imageio.plugins.png.PNGMetadata;
import org.bytedeco.javacpp.BytePointer;
import org.bytedeco.opencv.opencv_core.Mat;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageInputStream;
import javax.imageio.stream.ImageOutputStream;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collections;
import java.util.Iterator;

import static org.bytedeco.opencv.global.opencv_imgcodecs.imencode;
import static org.bytedeco.opencv.global.opencv_imgcodecs.imread;

/**
 * 图片处理工具
 *
 * @author 小天
 * @date 2019/6/3 23:29
 */
public class ImageTool {

    /**
     * jpeg 文件魔数，第 0 位
     */
    public static final byte JPEG_MAGIC_CODE_0           = (byte) 0xFF;
    public static final byte JPEG_MAGIC_CODE_1           = (byte) 0xD8;
    /**
     * 水平和垂直方向的像素密度单位：无单位
     */
    public static final byte JPEG_UNIT_OF_DENSITIES_NONE = 0x00;
    /**
     * 水平和垂直方向的像素密度单位：点数/英寸
     */
    public static final byte JPEG_UNIT_OF_DENSITIES_INCH = 0x01;
    /**
     * 水平和垂直方向的像素密度单位：点数/厘米
     */
    public static final byte UJPEG_NIT_OF_DENSITIES_CM   = 0x02;

    public static final byte DPI_120 = 0x78;
    public static final byte DPI_96  = 0x60;
    public static final byte DPI_72  = 0x48;

    /**
     * png 图片 pHYs 块，像素密度单位 / 每米
     */
    public static final int PNG_pHYs_pixelsPerUnit = (int) PoiUnitTool.centimeterToPixel(100);

    /**
     * 读取图片文件
     *
     * @param imgFile 图片文件
     *
     * @return {@link BufferedImage}
     *
     * @throws IOException
     */
    public static BufferedImage readImage(File imgFile) throws IOException {
        InputStream inputStream;
        if ((inputStream = FileUtil.readFile(imgFile)) == null) {
            return null;
        }
        return ImageIO.read(inputStream);
    }

    public static boolean isJpeg(BytePointer bytePointer) {
        return bytePointer.get(0) == JPEG_MAGIC_CODE_0 && bytePointer.get(1) == JPEG_MAGIC_CODE_1;
    }

    /**
     * 是否指定了正确的像素密度单位（仅适用于 jpeg）：
     * <ul>
     * <li>密度单位：点数/英寸</li>
     * <li>水平和垂直方向的像素密度：120</li>
     * </ul>
     */
    public static boolean isRightUnitOfDensities(BytePointer bytePointer) {
        return bytePointer.get(13) == JPEG_UNIT_OF_DENSITIES_INCH
                && bytePointer.get(14) == 0x00 && bytePointer.get(15) == DPI_120
                && bytePointer.get(16) == 0x00 && bytePointer.get(17) == DPI_120;
    }

    /**
     * 重置像素密度
     */
    public static void resetDensity(BytePointer bytePointer) {
        // 设置水平和垂直方向密度单位：1 点数/英寸
        bytePointer.put(13, JPEG_UNIT_OF_DENSITIES_INCH);
        // 设置水平方向像素密度：120
        bytePointer.put(14, (byte) 0);
        bytePointer.put(15, DPI_96);
        // 设置垂直方向像素密度：120
        bytePointer.put(16, (byte) 0);
        bytePointer.put(17, DPI_96);
    }

    /**
     * 图片转换为 jpeg
     */
    public static File convertToJpeg(File imgFile) throws IOException {
        Mat srcMat = imread(imgFile.getAbsolutePath());

        BytePointer srcBytePointer = srcMat.data();
        if (isJpeg(srcBytePointer)) {
            if (isRightUnitOfDensities(srcBytePointer)) {
                // 原样返回
                return imgFile;
            }
        } else {
            srcBytePointer = new BytePointer();
            // 转换为 jpg
            imencode(".jpg", srcMat, srcBytePointer);
        }

        // 重置像素密度信息
        resetDensity(srcBytePointer);

        // 生成新的图片文件
        File dstImgFile = TempFileUtil.createTempFile("jpg");
        try (OutputStream outputStream = new FileOutputStream(dstImgFile)) {
            outputStream.write(srcBytePointer.getStringBytes());
        }
        return dstImgFile;
    }

    /**
     * 重置 png 图片的 pHYs 信息，以修复 wps 在 win10 下打印图片缺失的 bug
     *
     * @param imageFile 源文件
     *
     * @return 新的文件，null：处理失败
     *
     * @throws IOException
     */
    public static File resetPhysOfPNG(File imageFile) throws IOException {
        ImageInputStream imageInputStream = ImageIO.createImageInputStream(imageFile);

        ImageReader reader = getImageReader(imageInputStream);
        if (reader == null) {
            return null;
        }
        if (!(reader instanceof PNGImageReader)) {
            return null;
        }
        reader.setInput(imageInputStream, true, false);
        BufferedImage bufferedImage;
        try {
            bufferedImage = reader.read(0, reader.getDefaultReadParam());
        } finally {
            reader.dispose();
            imageInputStream.close();
        }

        PNGMetadata metadata = (PNGMetadata) reader.getImageMetadata(0);
        if (metadata.pHYs_unitSpecifier != 1) {
            metadata.pHYs_pixelsPerUnitXAxis = PNG_pHYs_pixelsPerUnit;
            metadata.pHYs_pixelsPerUnitYAxis = PNG_pHYs_pixelsPerUnit;
            metadata.pHYs_unitSpecifier = 1;
            metadata.pHYs_present = true;
        }

        ImageOutputStream imageOutputStream = null;
        ImageWriter imageWriter = null;
        try {
            File dstImgFile = TempFileUtil.createTempFile("jpg");

            imageOutputStream = ImageIO.createImageOutputStream(dstImgFile);

            imageWriter = ImageIO.getImageWriter(reader);
            imageWriter.setOutput(imageOutputStream);
            imageWriter.write(new IIOImage(bufferedImage, Collections.emptyList(), metadata));

            return dstImgFile;
        } finally {
            if (imageWriter != null) {
                imageWriter.dispose();
            }
            if (imageWriter != null) {
                imageOutputStream.flush();
            }
        }
    }

    private static ImageReader getImageReader(ImageInputStream stream) {
        Iterator iter = ImageIO.getImageReaders(stream);
        if (!iter.hasNext()) {
            return null;
        }
        return (ImageReader) iter.next();
    }

}
