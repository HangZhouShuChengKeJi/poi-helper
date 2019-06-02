package com.orange.poi.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;

import java.io.File;
import java.io.IOException;
import java.util.Date;

/**
 * @author 小天
 * @date 2019/5/31 21:32
 */
public class TempFileUtil {

    private static File tempRootDir;

    static {
        tempRootDir = new File(System.getProperty("java.io.tmpdir"));
    }

    /**
     * 获取一个新的临时文件目录
     *
     * @param exName 文件扩展名
     *
     * @return 临时文件
     *
     * @throws IOException
     */
    public static File createTempFile(String exName) throws IOException {
        File tempFile;
        do {
            tempFile = new File(tempRootDir, generateNewName(new Date(), exName));
        } while (tempFile.exists());

        if (tempFile.createNewFile()) {
            return tempFile;
        }
        throw new IOException("创建临时文件失败：" + tempFile.getAbsolutePath());
    }

    private static String generateNewName(Date baseTime, String exName) {
        if (baseTime == null) {
            baseTime = new Date();
        }
        if (StringUtils.isNotBlank(exName)) {
            return DateFormatUtils.format(baseTime, "yyyyMMddHHmmssSSS")
                    + "_" + Thread.currentThread().getId()
                    + "_" + RandomStringUtils.random(8, true, true)
                    + "." + exName;
        }
        return DateFormatUtils.format(baseTime, "yyyyMMddHHmmssSSS")
                + "_" + Thread.currentThread().getId()
                + "_" + RandomStringUtils.random(8, true, true);
    }
}
