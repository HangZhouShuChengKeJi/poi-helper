package com.orange.poi.util;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;

import static org.junit.Assert.*;

/**
 * @author 小天
 * @date 2019/12/23 18:33
 */
public class FileTypeToolTest {

    @Before
    public void setUp() throws Exception {
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void detect() throws URISyntaxException, IOException {
        FileTypeTool fileTypeTool = FileTypeTool.getInstance();
        Assert.assertEquals(FileTypeEnum.PNG, fileTypeTool.detect(new File(getClass().getResource("/img/1_png.jpg").toURI())));
        Assert.assertEquals(FileTypeEnum.PNG, fileTypeTool.detect(new File(getClass().getResource("/img/1_png.png").toURI())));
        Assert.assertEquals(FileTypeEnum.PNG, fileTypeTool.detect(new File(getClass().getResource("/img/2_png.jpg").toURI())));
        Assert.assertEquals(FileTypeEnum.PNG, fileTypeTool.detect(new File(getClass().getResource("/img/2_png.png").toURI())));
        Assert.assertEquals(FileTypeEnum.JPEG, fileTypeTool.detect(new File(getClass().getResource("/img/3_jpg.jpg").toURI())));
        Assert.assertEquals(FileTypeEnum.JPEG, fileTypeTool.detect(new File(getClass().getResource("/img/3_jpg.png").toURI())));
    }
}