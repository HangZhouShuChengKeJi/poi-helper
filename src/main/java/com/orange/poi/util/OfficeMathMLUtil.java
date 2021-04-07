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
    private Transformer        mml2ommlTransformer;
    private Transformer        omml2mmlTransformer;

    private OfficeMathMLUtil() throws TransformerConfigurationException {
        transformerFactory = TransformerFactory.newInstance();

        StreamSource streamSource1 = new StreamSource(OfficeMathMLUtil.class.getClassLoader().getResourceAsStream("MML2OMML.XSL"));
        mml2ommlTransformer = transformerFactory.newTransformer(streamSource1);

        StreamSource streamSource2 = new StreamSource(OfficeMathMLUtil.class.getClassLoader().getResourceAsStream("OMML2MML.XSL"));
        omml2mmlTransformer = transformerFactory.newTransformer(streamSource2);
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
        mml2ommlTransformer.transform(srcDocSource, result);

        Document transformedDoc = result.getDocument();
        return transformedDoc.asXML();
    }

    /**
     * Office MathML 转换为 MathML
     *
     * @param omml Office MathML
     *
     * @return Office MathML
     *
     * @throws DocumentException
     * @throws TransformerException
     */
    public String convertOmmlToMml(String omml) throws DocumentException, TransformerException {

        Document srcDoc = DocumentHelper.parseText(omml);

        DocumentSource srcDocSource = new DocumentSource(srcDoc);

        DocumentResult result = new DocumentResult();
        omml2mmlTransformer.transform(srcDocSource, result);

        Document transformedDoc = result.getDocument();
        return transformedDoc.asXML();
    }
}
