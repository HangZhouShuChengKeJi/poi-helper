package com.orange.poi.util;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.io.DocumentResult;
import org.dom4j.io.DocumentSource;

import javax.xml.transform.Transformer;
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

    public static String MML2OMML_XSL;

    static {
        MML2OMML_XSL = OfficeMathMLUtil.class.getResource("/MML2OMML.XSl").getFile();
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
    public static String convertMmlToOmml(String mathML) throws DocumentException, TransformerException {

        Document sourceDoc = DocumentHelper.parseText(mathML);

        // todo 检查是否包含 namespace： http://www.w3.org/1998/Math/MathML

        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer transformer = factory.newTransformer(new StreamSource(MML2OMML_XSL));

        DocumentSource sourceDocSource = new DocumentSource(sourceDoc);
        DocumentResult result = new DocumentResult();
        transformer.transform(sourceDocSource, result);

        Document transformedDoc = result.getDocument();
        return transformedDoc.asXML();
    }
}
