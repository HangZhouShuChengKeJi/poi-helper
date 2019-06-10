package com.orange.poi.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * 文件工具类
 *
 * @author 小天
 * @date 2019/6/10 11:09
 */
public class FileUtil {

    /**
     * 读取文件，获取输入流
     *
     * @param imgFile 文件
     *
     * @return 文件输入流
     *
     * @throws FileNotFoundException
     */
    public static InputStream readFile(File imgFile) throws FileNotFoundException {
        String imgFilePath = imgFile.getAbsolutePath();
        int p = imgFilePath.indexOf(".jar!");
        if (p > 0) {
            // 对于 jar 包里的图片文件作特殊处理
            imgFilePath = imgFilePath.substring(p + 5).replace("\\", "/");
            return Thread.currentThread().getContextClassLoader().getResourceAsStream(imgFilePath);
        }
        return new FileInputStream(imgFile);
    }
}
