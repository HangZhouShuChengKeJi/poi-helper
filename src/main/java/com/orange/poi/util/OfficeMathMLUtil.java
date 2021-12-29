package com.orange.poi.util;

import org.apache.commons.collections4.CollectionUtils;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.XPath;
import org.dom4j.io.DocumentResult;
import org.dom4j.io.DocumentSource;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamSource;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * office mathML 工具类
 *
 * @author 小天
 * @date 2020/6/21 11:03
 */
public class OfficeMathMLUtil {


    private TransformerFactory transformerFactory;
    private Transformer mml2ommlTransformer;
    private Transformer omml2mmlTransformer;


    private Map<String, String> uriMap = new HashMap<>();


    private OfficeMathMLUtil() throws TransformerConfigurationException {
        transformerFactory = TransformerFactory.newInstance();

        StreamSource streamSource1 = new StreamSource(OfficeMathMLUtil.class.getClassLoader().getResourceAsStream("MML2OMML.XSL"));
        mml2ommlTransformer = transformerFactory.newTransformer(streamSource1);

        StreamSource streamSource2 = new StreamSource(OfficeMathMLUtil.class.getClassLoader().getResourceAsStream("OMML2MML.XSL"));
        omml2mmlTransformer = transformerFactory.newTransformer(streamSource2);

        uriMap.put("mml", "http://www.w3.org/1998/Math/MathML");
    }

    /**
     * 单例持有器
     */
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

    /**
     * 获取单例
     *
     * @return {@link OfficeMathMLUtil} 单例
     */
    public static OfficeMathMLUtil getInstance() {
        return SingleInstanceHolder.instance;
    }

    /**
     * 转换 mathml 里的 mfenced（主要用于方程组或矩阵）。
     *
     * mathjax 将 latex 的方程组或矩阵转换为 mathml 时，使用的是 mathml 的 mo 节点。
     * 但是 office 提供的 xsl 里仅支持过时的 mathml mfenced 节点。所以需要通过以下转换进行适配。
     *
     * @param mmlDoc
     */
    private void transformMfencedNode(Document mmlDoc) {
        // 创建 xPath
        XPath xPath = mmlDoc.createXPath("//mml:mo[@data-mjx-texclass='OPEN']");
        // xpath 里必须放入 uriMap
        xPath.setNamespaceURIs(uriMap);

        List<Node> nodeList = xPath.selectNodes(mmlDoc);

        if (CollectionUtils.isEmpty(nodeList)) {
            return;
        }
        for (Node node : nodeList) {
            Element parentEle = node.getParent();
            if (parentEle == null) {
                continue;
            }
            List<Element> childElementList = parentEle.elements();
            if (CollectionUtils.isEmpty(childElementList)) {
                continue;
            }
            Element firstEle = childElementList.get(0);
            if ("mo".equalsIgnoreCase(firstEle.getName())
                    && firstEle.attribute("data-mjx-texclass").getValue().equalsIgnoreCase("open")) {
                Element secondEle = childElementList.get(1);
                if ("mtable".equalsIgnoreCase(secondEle.getName())) {
                    Element thirdEle = childElementList.get(2);
                    if ("mo".equalsIgnoreCase(thirdEle.getName())
                            && thirdEle.attribute("data-mjx-texclass").getValue().equalsIgnoreCase("close")) {

                        Element mfencedEle = parentEle.addElement("mfenced");
                        mfencedEle.addAttribute("open", firstEle.getText());
                        mfencedEle.addAttribute("close", thirdEle.getText());
                        mfencedEle.add(secondEle.createCopy());
                        parentEle.remove(firstEle);
                        parentEle.remove(secondEle);
                        parentEle.remove(thirdEle);
                    }
                }
            }
        }
    }

    /**
     * 预处理 mathml
     *
     * @param mmlDoc
     */
    private void preProcessMathML(Document mmlDoc) {
        transformMfencedNode(mmlDoc);
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

        // mathml 对象预处理
        preProcessMathML(srcDoc);

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
