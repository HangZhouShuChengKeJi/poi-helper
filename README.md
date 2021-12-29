# poi-helper
[Apache poi](https://poi.apache.org/) 库辅助工具，提供更便捷的操作 word、excel 等的方法

# 使用
在 maven 工程里，添加依赖：
```xml
    <dependency>
        <groupId>com.orange.opensource</groupId>
        <artifactId>poi-helper</artifactId>
        <version>0.0.4-SNAPSHOT</version>
    </dependency>
```

示例代码：
```java
package com.orange.poi.test;

import com.orange.poi.util.TempFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
* 
* 测试
* 
* @author 小天
* @date 2019/6/5 10:59
*/
public class App {
    
    private String defaultFontFamily = "思源宋体";
    private int    defaultFontSize   = 14;
    private String defaultColor      = "000000";
    
    public static void main(String[] args) throws IOException {
         // 创建文档
         XWPFDocument doc = new XWPFDocument();
         
         // 初始化为 A4 纸大小
         PoiWordTool.initDocForA4(doc);
        
         // 创建段落
         XWPFParagraph paragraph = doc.createParagraph();
         
         // 添加文字
         PoiWordParagraphTool.addTxt(paragraph, "hello world", defaultFontFamily, defaultFontSize, defaultColor);
        
         // 创建输出文件
         File wordFile = TempFileUtil.createTempFile("docx");
         
         // 显示输出的文件路径
         System.out.println(wordFile);
        
         // 输出到文件
         FileOutputStream out = new FileOutputStream(wordFile);
         doc.write(out);
         out.close();
    }
}
```

# 参考
+ [Office Open XML](http://www.officeopenxml.com/)
+ [ECMA-376 | Office Open XML file formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
+ [ISO/IEC 29500-1:2016 | Office Open XML File Formats — Part 1: Fundamentals and Markup Language Reference](https://www.iso.org/standard/71691.html)
+ [ISO/IEC 29500-2:2021 | Office Open XML file formats — Part 2: Open packaging conventions](https://www.iso.org/standard/77818.html)
+ [ISO/IEC 29500-3:2015 | Office Open XML File Formats — Part 3: Markup Compatibility and Extensibility](https://www.iso.org/standard/65533.html)
+ [ISO/IEC 29500-4:2016 | Office Open XML File Formats — Part 4: Transitional Migration Features](https://www.iso.org/standard/71692.html)
+ [apache poi 库](https://poi.apache.org/)
+ [MathML | MDN 的 mathml 文档](https://developer.mozilla.org/en-US/docs/Web/MathML)

# 版权

Licensed under the [Apache 2.0 license](https://www.apache.org/licenses/LICENSE-2.0.html) license.

