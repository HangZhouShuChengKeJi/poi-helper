package com.orange.poi.util;

import org.dom4j.DocumentException;
import org.junit.Test;

import javax.xml.transform.TransformerException;

/**
 *
 * @author 小天
 * @date 2020/6/21 11:14
 */
public class OfficeMathMLUtilTest {

    @Test
    public void convertMmlToOmml() {
        String text = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                "<mfrac>" +
                "<mi>a</mi>" +
                "<mi>b</mi>" +
                "</mfrac>" +
                "</math>";
        try {
            System.out.println(OfficeMathMLUtil.getInstance().convertMmlToOmml(text));
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        }
    }
}