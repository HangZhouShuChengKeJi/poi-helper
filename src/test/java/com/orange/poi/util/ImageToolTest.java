package com.orange.poi.util;

import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;

/**
 * @author cjie
 * @date 2021/4/8.
 */
public class ImageToolTest {

    @Test
    public void resetDensityTest() throws URISyntaxException, IOException {
        File img = new File(getClass().getResource("/img/9.jpg").toURI());
        ImageTool.resetDensity(img);
    }
}
