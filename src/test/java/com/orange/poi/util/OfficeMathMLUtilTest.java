package com.orange.poi.util;

import org.apache.commons.io.IOUtils;
import org.dom4j.DocumentException;
import org.junit.Test;

import javax.xml.transform.TransformerException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

/**
 *
 * @author 小天
 * @date 2020/6/21 11:14
 */
public class OfficeMathMLUtilTest {

    @Test
    public void convertMmlToOmml() throws IOException, DocumentException, TransformerException {
        String mml;
        String omml;

        mml = IOUtils.resourceToString("/mml/mml-frac.xml", StandardCharsets.UTF_8);
        omml = OfficeMathMLUtil.getInstance().convertMmlToOmml(mml);
        System.out.println(omml);

        mml = IOUtils.resourceToString("/mml/mml-sqrt.xml", StandardCharsets.UTF_8);
        omml = OfficeMathMLUtil.getInstance().convertMmlToOmml(mml);
        System.out.println(omml);

        mml = IOUtils.resourceToString("/mml/mml-mfenced.xml", StandardCharsets.UTF_8);
        omml = OfficeMathMLUtil.getInstance().convertMmlToOmml(mml);
        System.out.println(omml);

        mml = IOUtils.resourceToString("/mml/mml-env-cases.xml", StandardCharsets.UTF_8);
        omml = OfficeMathMLUtil.getInstance().convertMmlToOmml(mml);
        System.out.println(omml);
    }
}