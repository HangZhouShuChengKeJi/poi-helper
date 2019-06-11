package com.orange.poi.util;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * 图片处理工具
 *
 * @author 小天
 * @date 2019/6/3 23:29
 */
public class ImageTool {

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

    /**
     * 等比例缩放图片
     *
     * @param imgFile   图片文件
     * @param maxWidth  缩放后的最大宽度，单位：像素
     * @param maxHeight 缩放后的最大高度，单位：像素
     * @param lockScale 锁定缩放比例
     *
     * @return 图片信息 {@link ImageInfo}
     *
     * @throws IOException
     */
    public static ImageInfo resizeImage(File imgFile, final int maxWidth, final int maxHeight, boolean lockScale) throws IOException {
        // todo 实现不按照比例重绘功能
        final BufferedImage image = readImage(imgFile);
        if (image == null) {
            throw new IllegalArgumentException("图片文件不存在： " + imgFile.getAbsolutePath());
        }
        final int actualWidth = image.getWidth();
        final int actualHeight = image.getHeight();
        if (actualWidth > maxWidth || actualHeight > maxHeight) {
            // 注意计算过程的精度问题
            final double scaleW = (double) actualWidth / maxWidth;
            final double scaleH = (double) actualHeight / maxHeight;
            final int newWidth;
            final int newHeight;
            if (scaleW > scaleH) {
                newWidth = maxWidth;
                newHeight = (int) (actualHeight / scaleW);
            } else {
                newWidth = (int) (actualWidth / scaleH);
                newHeight = maxHeight;
            }

            final String imgEx = UrlUtil.getExNameFromUrl(imgFile.getAbsolutePath());

            final BufferedImage newImage = new BufferedImage(newWidth, newHeight, BufferedImage.TYPE_INT_RGB);
            newImage.getGraphics().drawImage(image, 0, 0, newWidth, newHeight, null);

            File newImgFile = TempFileUtil.createTempFile(imgEx);

            try (FileOutputStream fileOutputStream = new FileOutputStream(newImgFile)) {
                ImageIO.write(newImage, imgEx, fileOutputStream);
                return new ImageInfo(newImgFile, newWidth, newHeight);
            }
        }
        // 返回原始数据
        return new ImageInfo(imgFile, actualWidth, actualHeight);
    }

    /**
     * 图片信息
     */
    public static class ImageInfo {
        private File imgFile;
        private int  width;
        private int  height;

        public ImageInfo(File imgFile, int width, int height) {
            this.imgFile = imgFile;
            this.width = width;
            this.height = height;
        }

        public File getImgFile() {
            return imgFile;
        }

        public int getWidth() {
            return width;
        }

        public int getHeight() {
            return height;
        }
    }
}
