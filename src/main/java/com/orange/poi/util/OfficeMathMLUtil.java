package com.orange.poi.util;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.io.DocumentResult;
import org.dom4j.io.DocumentSource;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamSource;

/**
 * office mathML 工具类
 *
 * @author 小天
 * @date 2020/6/21 11:03
 */
public class OfficeMathMLUtil {


    private TransformerFactory transformerFactory;
    private Transformer transformer;

    private OfficeMathMLUtil() throws TransformerConfigurationException {
        transformerFactory = TransformerFactory.newInstance();

        StreamSource streamSource = new StreamSource(OfficeMathMLUtil.class.getClassLoader().getResourceAsStream("MML2OMML.XSL"));
        transformer = transformerFactory.newTransformer(streamSource);
    }

    private static class SingleInstanceHolder {
        private static volatile OfficeMathMLUtil instance;
        static {
            try {
                instance = new OfficeMathMLUtil();
            } catch (TransformerConfigurationException e) {
                throw new RuntimeException(e);
            }
        }
    }

    public static OfficeMathMLUtil getInstance() {
        return SingleInstanceHolder.instance;
    }

    /**
     * MathML 转换为 Office MathML
     *
     * @param mathML MathML
     *
     * @return Office MathML
     *
     * @throws DocumentException
     * @throws TransformerException
     */
    public String convertMmlToOmml(String mathML) throws DocumentException, TransformerException {

        Document srcDoc = DocumentHelper.parseText(mathML);

        DocumentSource srcDocSource = new DocumentSource(srcDoc);

        DocumentResult result = new DocumentResult();
        transformer.transform(srcDocSource, result);

        Document transformedDoc = result.getDocument();
        return transformedDoc.asXML();
    }
}
