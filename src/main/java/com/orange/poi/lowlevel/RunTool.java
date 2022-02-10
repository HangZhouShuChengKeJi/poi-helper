package com.orange.poi.lowlevel;

import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;

/**
 * @author 小天
 * @date 2022/2/10 11:54
 */
public class RunTool {

    /**
     * 获取域代码类型
     *
     * @param ctr 段落 {@link CTR}
     *
     * @return 域代码类型 {@link STFldCharType.Enum}
     */
    public static STFldCharType.Enum getFldCharType(CTR ctr) {
        CTFldChar[] fldCharArr = ctr.getFldCharArray();
        if (fldCharArr.length > 0) {
            CTFldChar fldChar = fldCharArr[0];
            return fldChar.getFldCharType();
        }
        return null;
    }

    /**
     * 获取 instrText（域代码）内容
     *
     * @param ctr 段落 {@link CTR}
     *
     * @return instrText（域代码）内容
     */
    public static String getInstrTxt(CTR ctr) {
        CTText[] ctTexts = ctr.getInstrTextArray();
        if (ctTexts.length > 0) {
            return ctTexts[0].getStringValue().trim();
        }
        return null;
    }
}
