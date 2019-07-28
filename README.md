# poi-helper
[Apache poi](https://poi.apache.org/) 库辅助工具，提供更便捷的操作 word、excel 等的方法

# 使用
在 maven 工程里，添加依赖：
```xml
    <dependency>
        <groupId>com.orange.opensource</groupId>
        <artifactId>poi-helper</artifactId>
        <version>${eventframework.version}</version>
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

# 版权
[Apache License, Version 2.0](http://www.apache.org/licenses/LICENSE-2.0.html) Copyright (C) [杭州数橙科技有限公司](https://github.com/HangZhouShuChengKeJi)

