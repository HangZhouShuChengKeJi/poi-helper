package com.orange.poi.util;

import com.orange.poi.PoiUnitTool;
import com.sun.imageio.plugins.jpeg.JPEG;
import com.sun.imageio.plugins.jpeg.JPEGImageReader;
import com.sun.imageio.plugins.jpeg.JPEGMetadata;
import com.sun.imageio.plugins.png.PNGMetadata;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;

import javax.imageio.IIOException;
import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.ImageWriteParam;
import javax.imageio.ImageWriter;
import javax.imageio.metadata.IIOInvalidTreeException;
import javax.imageio.metadata.IIOMetadata;
import javax.imageio.metadata.IIOMetadataNode;
import javax.imageio.plugins.jpeg.JPEGImageWriteParam;
import javax.imageio.stream.ImageInputStream;
import javax.imageio.stream.ImageOutputStream;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collections;
import java.util.Iterator;

/**
 * 图片处理工具
 *
 * @author 小天
 * @date 2019/6/3 23:29
 */
public class ImageTool {

    private final static Logger logger = LoggerFactory.getLogger(ImageTool.class);

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
    public static final byte JPEG_NIT_OF_DENSITIES_CM    = 0x02;

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

    /**
     * 重置图片的像素密度信息（默认重置为 96，只支持 jpg 和 png 图片），以修复 wps 在 win10 下打印图片缺失的 bug
     *
     * @param imageFile 源文件
     *
     * @return 新的文件，null：处理失败
     *
     * @throws IOException
     */
    public static File resetDensity(File imageFile) throws IOException {
        ImageInputStream imageInputStream = ImageIO.createImageInputStream(imageFile);
        if (imageInputStream == null) {
            return null;
        }

        ImageReader reader = getImageReader(imageInputStream);
        if (reader == null) {
            return null;
        }

        reader.setInput(imageInputStream, true, false);

        String exName;
        IIOMetadata metadata;
        try {
            metadata = reader.getImageMetadata(0);
        } catch (IIOException e) {
            logger.error("imageFile={}", imageFile, e);
            return null;
        }
        if (metadata instanceof JPEGMetadata) {
            JPEGMetadata jpegMetadata = (JPEGMetadata) metadata;
            Integer resUnits = getResUnits(jpegMetadata);
            if (resUnits != null && resUnits != 0) {
                // 已指定了像素密度时，不再继续处理
                return null;
            }
            resetDensity(jpegMetadata);
            exName = "jpg";
        } else if (metadata instanceof PNGMetadata) {
            PNGMetadata pngMetadata = (PNGMetadata) metadata;
            if (pngMetadata.pHYs_unitSpecifier != 0) {
                // 已指定了像素密度时，不再继续处理
                return null;
            }
            resetDensity(pngMetadata);
            exName = "png";
        } else {
            throw new IllegalArgumentException("不支持的图片格式");
        }

        BufferedImage bufferedImage;
        try {
            bufferedImage = reader.read(0, reader.getDefaultReadParam());
        } finally {
            reader.dispose();
            imageInputStream.close();
        }

        ImageOutputStream imageOutputStream = null;
        ImageWriter imageWriter = null;
        try {
            File dstImgFile = TempFileUtil.createTempFile(exName);

            imageOutputStream = ImageIO.createImageOutputStream(dstImgFile);

            imageWriter = ImageIO.getImageWriter(reader);
            imageWriter.setOutput(imageOutputStream);

            ImageWriteParam writeParam = imageWriter.getDefaultWriteParam();
            if (writeParam instanceof JPEGImageWriteParam) {
                ((JPEGImageWriteParam) writeParam).setOptimizeHuffmanTables(true);
            }
            try {
                imageWriter.write(metadata, new IIOImage(bufferedImage, Collections.emptyList(), metadata), writeParam);
            } catch (NullPointerException e) {
                //有些时候会出现LCMS.getProfileSize出现空指针，因为LCMS单例被注销  原因不明 也无法避免  故try catch下
                logger.error("imageFile={}", imageFile, e);
                return null;
            }
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

    private static void resetDensity(JPEGMetadata metadata) throws IIOInvalidTreeException {
        final IIOMetadataNode newRootNode = new IIOMetadataNode(JPEG.nativeImageMetadataFormatName);

        // 方法一
        final IIOMetadataNode mergeJFIFsubNode = new IIOMetadataNode("mergeJFIFsubNode");
        IIOMetadataNode jfifNode = new IIOMetadataNode("jfif");
        jfifNode.setAttribute("majorVersion", null);
        jfifNode.setAttribute("minorVersion", null);
        jfifNode.setAttribute("thumbWidth", null);
        jfifNode.setAttribute("thumbHeight", null);

        // 重置像素密度单位
        jfifNode.setAttribute("resUnits", "1");
        jfifNode.setAttribute("Xdensity", "96");
        jfifNode.setAttribute("Ydensity", "96");
        mergeJFIFsubNode.appendChild(jfifNode);

        newRootNode.appendChild(mergeJFIFsubNode);
        newRootNode.appendChild(new IIOMetadataNode("mergeSequenceSubNode"));
        metadata.mergeTree(JPEG.nativeImageMetadataFormatName, newRootNode);

        // 方法二
//        final IIOMetadataNode dimensionNode = new IIOMetadataNode("Dimension");
//        final IIOMetadataNode horizontalPixelSizeNode = new IIOMetadataNode("HorizontalPixelSize");
//        horizontalPixelSizeNode.setAttribute("value", String.valueOf(25.4f / 96));
//        final IIOMetadataNode verticalPixelSizeNode = new IIOMetadataNode("VerticalPixelSize");
//        verticalPixelSizeNode.setAttribute("value", String.valueOf(25.4f / 96));
//        dimensionNode.appendChild(horizontalPixelSizeNode);
//        dimensionNode.appendChild(verticalPixelSizeNode);
//        newRootNode.appendChild(dimensionNode);
//        metadata.mergeTree(IIOMetadataFormatImpl.standardMetadataFormatName, newRootNode);
    }

    private static void resetDensity(PNGMetadata metadata) throws IIOInvalidTreeException {
        metadata.pHYs_pixelsPerUnitXAxis = PNG_pHYs_pixelsPerUnit;
        metadata.pHYs_pixelsPerUnitYAxis = PNG_pHYs_pixelsPerUnit;
        metadata.pHYs_unitSpecifier = 1;
        metadata.pHYs_present = true;
    }

    /**
     * 获取 jpg 图片的像素密度类型
     *
     * @param metadata
     *
     * @return
     */
    private static Integer getResUnits(JPEGMetadata metadata) {
        String value = getJfifAttr(metadata, "resUnits");
        if (value == null) {
            return null;
        }
        return Integer.parseInt(value);
    }

    /**
     * 获取 jpg 图片的像素密度类型
     *
     * @param metadata
     *
     * @return
     */
    private static String getJfifAttr(JPEGMetadata metadata, String attrName) {
        Node metadataNode = metadata.getAsTree(JPEG.nativeImageMetadataFormatName);

        if (metadataNode != null) {
            Node child = metadataNode.getFirstChild();
            while (child != null) {
                if (child.getNodeName().equals("JPEGvariety")) {
                    Node subChild = child.getFirstChild();
                    while (subChild != null) {
                        if ("app0JFIF".equals(subChild.getNodeName())) {
                            Node valueNode = subChild.getAttributes().getNamedItem(attrName);
                            if (valueNode != null) {
                                return valueNode.getNodeValue();
                            }
                            break;
                        }
                        subChild = subChild.getNextSibling();
                    }
                    break;
                }
                child = child.getNextSibling();
            }
        }
        return null;
    }

    private static ImageReader getImageReader(ImageInputStream stream) {
        Iterator iter = ImageIO.getImageReaders(stream);
        if (!iter.hasNext()) {
            return null;
        }
        return (ImageReader) iter.next();
    }

    /**
     * 获取像素密度
     *
     * @param imageFile 源文件
     *
     * @return 像素密度
     *
     * @throws IOException
     */
    public static Integer getDensity(File imageFile) throws IOException {
        ImageInputStream imageInputStream = ImageIO.createImageInputStream(imageFile);
        if (imageInputStream == null) {
            return null;
        }

        ImageReader reader = getImageReader(imageInputStream);
        if (reader == null) {
            return null;
        }
        if (!(reader instanceof JPEGImageReader)) {
            return null;
        }
        reader.setInput(imageInputStream, true, false);

        IIOMetadata metadata;
        try {
            metadata = reader.getImageMetadata(0);
        } catch (IIOException e) {
            logger.error("imageFile={}", imageFile, e);
            return null;
        }
        if (metadata instanceof JPEGMetadata) {
            JPEGMetadata jpegMetadata = (JPEGMetadata) metadata;
            Integer resUnits = getResUnits(jpegMetadata);
            if (resUnits == null) {
                return null;
            }
            if (resUnits == 1) {
                // 暂时只支持 resUnits == 1 等情况
                String value = getJfifAttr(jpegMetadata, "Xdensity");
                if (value == null) {
                    return null;
                }
                return Integer.parseInt(value);
            }
            return null;
        } else if (metadata instanceof PNGMetadata) {
            PNGMetadata pngMetadata = (PNGMetadata) metadata;
            if (pngMetadata.pHYs_unitSpecifier == 1) {
                // 暂时只支持 pHYs_unitSpecifier == 1 等情况
                return pngMetadata.pHYs_pixelsPerUnitXAxis;
            }
            return null;
        } else {
            throw new IllegalArgumentException("不支持的图片格式");
        }
    }

}
