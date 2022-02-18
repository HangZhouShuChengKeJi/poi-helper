package com.orange.poi.word;

import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author 小天
 * @date 2022/2/9 12:25
 */
public class PoiWordHeaderFooterToolTest {

    @Before
    public void setUp() throws Exception {
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void createHeaderFooterPolicy() {
    }

    @Test
    public void createHeader() {
    }

    @Test
    public void createFooter() {
    }

    @Test
    public void removeHeader() throws IOException {
        InputStream wordInputStream = getClass().getResourceAsStream("/word/word_001.docx");

        XWPFDocument wordDoc = new XWPFDocument(wordInputStream);

        // 移除除了第一节之外的所有页眉。 后续的节没有页眉页脚时，会自动使用第一节的页眉页脚。
        boolean find = false;
        for (XWPFParagraph paragraph : wordDoc.getParagraphs()) {
            CTSectPr sectPr = PoiWordSectionTool.getSectionProperties(paragraph);
            if (sectPr == null) {
                continue;
            }
            if (!find) {
                find = true;
                continue;
            }
            PoiWordHeaderFooterTool.removeHeader(sectPr);
        }

        File newWordFile = TempFileUtil.createTempFile("docx");
        System.out.println(newWordFile);
        FileOutputStream out = new FileOutputStream(newWordFile);
        wordDoc.write(out);
        out.close();
    }

    @Test
    public void removeFooter() throws IOException {
        InputStream wordInputStream = getClass().getResourceAsStream("/word/word_001.docx");

        XWPFDocument wordDoc = new XWPFDocument(wordInputStream);

        // 移除除了第一节之外的所有页脚。 后续的节没有页眉页脚时，会自动使用第一节的页眉页脚。
        boolean find = false;
        for (XWPFParagraph paragraph : wordDoc.getParagraphs()) {
            CTSectPr sectPr = PoiWordSectionTool.getSectionProperties(paragraph);
            if (sectPr == null) {
                continue;
            }
            if (!find) {
                find = true;
                continue;
            }
            PoiWordHeaderFooterTool.removeFooter(sectPr);
        }

        File newWordFile = TempFileUtil.createTempFile("docx");
        System.out.println(newWordFile);
        FileOutputStream out = new FileOutputStream(newWordFile);
        wordDoc.write(out);
        out.close();
    }
}