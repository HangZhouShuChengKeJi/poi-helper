package com.orange.poi.util;

import org.junit.Test;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.Iterator;

/**
 * @author cjie
 * @date 2021/4/8.
 */
public class ImageToolTest {

    @Test
    public void resetDensityTest() throws URISyntaxException, IOException {
//        File img = new File(getClass().getResource("/img/10.jpg").toURI());
        File img = new File(getClass().getResource("/img/11.png").toURI());
        File file = ImageTool.resetDensity(img);
    }
}
