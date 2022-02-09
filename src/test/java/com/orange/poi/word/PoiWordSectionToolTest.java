package com.orange.poi.word;

import com.orange.poi.PoiUnitTool;
import com.orange.poi.util.TempFileUtil;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;


/**
 * @author 小天
 * @date 2022/1/25 16:18
 */
public class PoiWordSectionToolTest {


    private String defaultFontFamily = "宋体";
    private int    defaultFontSize   = 14;
    private String defaultColor      = "000000";

    @Before
    public void setUp() throws Exception {
        System.setProperty("java.io.tmpdir", new File("temp").getAbsolutePath());
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void setCols() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        // 设置默认样式
        PoiWordTool.setDefaultStyle(doc, "Arial", "宋体", 14, "000000");
        // 设置默认分栏
        PoiWordSectionTool.setCols(doc, 2, 10, true);
        // 设置默认 section 不分页
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(doc), STSectionMark.NEXT_PAGE);

        // 第 1 节

        PoiWordParagraphTool.createParagraph(doc, "第 1 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        // 第 1 节的分节符
        XWPFParagraph sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        // 第 1 节 的 type 设置不设置效果都一样。
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.CONTINUOUS);

        // 第 2 节

        PoiWordParagraphTool.createParagraph(doc, "第 2 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        PoiWordParagraphTool.createParagraph(doc, "本节从新的页开始");
        // 第 2 节的分节符
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.NEXT_PAGE);

        // 第 3 节

        PoiWordParagraphTool.createParagraph(doc, "第 3 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        PoiWordParagraphTool.createParagraph(doc, "本节和前一节挨在一起");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.CONTINUOUS);

        // 第 4 节

        PoiWordParagraphTool.createParagraph(doc, "第 4 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        PoiWordParagraphTool.createParagraph(doc, "本节从新的页开始");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.NEXT_PAGE);

        // 第 5 节

        PoiWordParagraphTool.createParagraph(doc, "第 5 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        // 导出的 word 可能看不出来，可以转换位 pdf，就可以看到该效果
        PoiWordParagraphTool.createParagraph(doc, "本节从新的奇数页开始");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.ODD_PAGE);

        // 第 6 节

        PoiWordParagraphTool.createParagraph(doc, "第 6 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        // 导出的 word 可能看不出来，可以转换位 pdf，就可以看到该效果
        PoiWordParagraphTool.createParagraph(doc, "本节从新的奇数页开始");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.ODD_PAGE);

        // 最后一节 - 从新的页开始

        for (int i = 0; i < 2; i++) {
            XWPFParagraph paragraph = PoiWordParagraphTool.createParagraph(doc, String.format("第 %d 段。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。", i));

            PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.0,1.0);
        }

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }

    /**
     * 测试不同节、页眉页脚、页码等综合功能
     *
     * @throws IOException
     */
    @Test
    public void testHeaderFooter() throws IOException {
        XWPFDocument doc = PoiWordTool.createDocForA4();
        // 设置默认样式
        PoiWordTool.setDefaultStyle(doc, "Arial", "宋体", 14, "000000");
        // 设置默认分栏
        PoiWordSectionTool.setCols(doc, 2, 10, true);
        // 设置默认 section 不分页
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(doc), STSectionMark.NEXT_PAGE);

        // 设置奇偶页页眉页脚不同
        PoiWordHeaderFooterTool.setEvenAndOddHeaders(doc, true);
        // 页眉距离顶部的距离
        PoiWordHeaderFooterTool.setHeaderMargin(doc, PoiUnitTool.centimeterToDXA(1.500f));
        // 页脚距离底部的距离
        PoiWordHeaderFooterTool.setFooterMargin(doc, PoiUnitTool.centimeterToDXA(1.200f));

        // 页眉
        XWPFHeader xwpfHeader = PoiWordHeaderFooterTool.createHeader(doc, HeaderFooterType.DEFAULT);
        XWPFParagraph headerParagraph = xwpfHeader.createParagraph();
        PoiWordParagraphTool.addTxt(headerParagraph, "分节 | 页眉页脚 | 页码 | 测试");

        // 奇数页页脚
        XWPFFooter xwpfFooter = PoiWordHeaderFooterTool.createFooter(doc, HeaderFooterType.DEFAULT);
        XWPFParagraph footerParagraph = xwpfFooter.createParagraph();
        // 奇数页页脚的页码右对齐
        footerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        PoiWordParagraphTool.addTxt(footerParagraph, "第", defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordPageNoTool.addPageNo(footerParagraph, defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.addTxt(footerParagraph, "页", defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);

        // 偶数页页脚
        xwpfFooter = PoiWordHeaderFooterTool.createFooter(doc, HeaderFooterType.EVEN);
        footerParagraph = xwpfFooter.createParagraph();
        // 偶数页页脚的页码左对齐
        footerParagraph.setAlignment(ParagraphAlignment.LEFT);
        PoiWordParagraphTool.addTxt(footerParagraph, "第", defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordPageNoTool.addPageNo(footerParagraph, defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);
        PoiWordParagraphTool.addTxt(footerParagraph, "页", defaultFontFamily, defaultFontFamily, defaultFontSize, defaultColor);

        // 第 1 节

        PoiWordParagraphTool.createParagraph(doc, "第 1 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        // 第 1 节的分节符
        XWPFParagraph sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        // 第 1 节 的 type 设置不设置效果都一样。
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.CONTINUOUS);

        // 第 2 节

        PoiWordParagraphTool.createParagraph(doc, "第 2 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        PoiWordParagraphTool.createParagraph(doc, "本节从新的页开始");
        // 第 2 节的分节符
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.NEXT_PAGE);

        PoiWordParagraphTool.addPageBreak(doc);
        PoiWordParagraphTool.addPageBreak(doc);
        PoiWordParagraphTool.addPageBreak(doc);
        PoiWordParagraphTool.addPageBreak(doc);
        PoiWordParagraphTool.addPageBreak(doc);

        // 第 3 节

        PoiWordParagraphTool.createParagraph(doc, "第 3 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        PoiWordParagraphTool.createParagraph(doc, "本节和前一节挨在一起");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.CONTINUOUS);

        // 第 4 节

        PoiWordParagraphTool.createParagraph(doc, "第 4 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        PoiWordParagraphTool.createParagraph(doc, "本节从新的页开始");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.NEXT_PAGE);

        // 第 5 节

        PoiWordParagraphTool.createParagraph(doc, "第 5 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        // 导出的 word 可能看不出来，可以转换位 pdf，就可以看到该效果
        PoiWordParagraphTool.createParagraph(doc, "本节从新的奇数页开始");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.ODD_PAGE);

        // 第 6 节

        PoiWordParagraphTool.createParagraph(doc, "第 6 节 - 第一段 - 不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏不需要分栏");
        // 导出的 word 可能看不出来，可以转换位 pdf，就可以看到该效果
        PoiWordParagraphTool.createParagraph(doc, "本节从新的奇数页开始");
        sectionParagraph = PoiWordSectionTool.addSection(doc, false);
        PoiWordSectionTool.setType(PoiWordSectionTool.getSectionProperties(sectionParagraph), STSectionMark.ODD_PAGE);

        // 最后一节 - 从新的页开始

        for (int i = 0; i < 2; i++) {
            XWPFParagraph paragraph = PoiWordParagraphTool.createParagraph(doc, String.format("第 %d 段。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。" +
                    "段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。段落内容比较长。", i));

            PoiWordParagraphTool.setParagraphSpaceOfLine(paragraph, 1.0,1.0);
        }

        File wordFile = TempFileUtil.createTempFile("docx");

        System.out.println(wordFile);

        FileOutputStream out = new FileOutputStream(wordFile);
        doc.write(out);
        out.close();
    }
}