package com.orange.poi.word;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRelation;

/**
 * @author 小天
 * @date 2022/2/16 14:40
 */
public class PoiWordRelationTool {

    public static int getRelationIndex(XWPFDocument doc, XWPFRelation relation) {
        int i = 1;
        for (POIXMLDocumentPart.RelationPart rp : doc.getRelationParts()) {
            if (rp.getRelationship().getRelationshipType().equals(relation.getRelation())) {
                i++;
            }
        }
        return i;
    }
}
