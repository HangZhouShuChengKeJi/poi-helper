package com.orange.poi.util;

import org.apache.commons.codec.binary.Hex;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * 文件类型检测工具类（单例）
 *
 * @author 小天
 * @date 2019/12/23 16:53
 */
public class FileTypeTool {

    /**
     * 索引（{@link FileTypeEnum#magicCode}）最大长度
     */
    private Integer                   maxByteSize = 0;
    /**
     * {@link FileTypeEnum#magicCode} 长度分布菜单，降序排列
     */
    private List<Integer>             lengthMenu;
    /**
     * 文件类型索引 map<br/>
     * key： {@link FileTypeEnum#magicCode}<br/>
     * value： {@link FileTypeEnum}<br/>
     */
    private Map<String, FileTypeEnum> fileTypeEnumMap;

    private FileTypeTool() {
        fileTypeEnumMap = new TreeMap<>();
        FileTypeEnum[] fileTypeEnums = FileTypeEnum.values();
        Set<Integer> lengthSet = new HashSet<>();
        for (FileTypeEnum typeEnum : fileTypeEnums) {
            if (typeEnum.magicCode != null && typeEnum.magicCode.length() > 0) {
                fileTypeEnumMap.put(typeEnum.magicCode, typeEnum);

                int byteLen = typeEnum.magicCode.length() / 2;
                lengthSet.add(byteLen);
                if (byteLen > maxByteSize) {
                    maxByteSize = byteLen;
                }
            }
        }
        lengthMenu = new LinkedList<>(lengthSet);
        // 降序排列
        lengthMenu.sort(Collections.reverseOrder(Integer::compareTo));
    }

    /**
     * 使用静态内部类实现单例
     */
    private static class FileTypeToolInstance {
        private final static FileTypeTool instance = new FileTypeTool();
    }

    /**
     * 单例
     */
    public static FileTypeTool getInstance() {
        return FileTypeToolInstance.instance;
    }

    /**
     * 检测文件类型
     *
     * @param file 文件
     *
     * @return {@link FileTypeEnum}，检测不到时，返回 null
     */
    public FileTypeEnum detect(File file) {
        if (file == null) {
            return null;
        }
        if (!file.exists()) {
            return null;
        }
        if (!file.isFile()) {
            return null;
        }
        try (FileInputStream fileInputStream = new FileInputStream(file)){
            byte[] byteArr = new byte[maxByteSize];
            int readByteSize = fileInputStream.read(byteArr);
            if (readByteSize == -1) {
                throw new IllegalArgumentException("文件内容为空");
            }
            if (readByteSize < maxByteSize) {
                throw new IllegalArgumentException("文件内容格式错误");
            }
            String hexStr = Hex.encodeHexString(byteArr).toUpperCase();
            for (Integer length : lengthMenu) {
                FileTypeEnum fileTypeEnum = fileTypeEnumMap.get(hexStr.substring(0, length));
                if (fileTypeEnum != null) {
                    return fileTypeEnum;
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return null;
    }

}
