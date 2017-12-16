package org.lex.doc2others.convert.ppt;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Lex.Shaoli.Zhu
 *
 */
public class TestPPTConvert {

    private static Logger log = LoggerFactory.getLogger(TestPPTConvert.class);
    private static final String ORIGIN_PPT_PATH = "src/main/resources/ppt/test.pptx";
    // private static final String TARGET_IMAGES_PATH =
    // "src/main/resources/ppt/images";
    private static final boolean isVisble = true;

    @Test
    public void ppt2ImageTest() throws Exception {
        log.info("----实例化PPTConvert, 开始...");
        PPTConvertor pptConvert = new PPTConvertor(ORIGIN_PPT_PATH, isVisble);
        // log.info("----test saveTo : {}", pptConvert.getSavetoPath(new
        // File(ORIGIN_PPT_PATH)));
        log.info("----实例化PPTConvert, 成功...");
        log.info("----PPT开始转换为Image");
        long s = System.currentTimeMillis();
        // pptConvert.PPTToJPG(PPTConvertor.SaveToType.PPT_SAVEAS_);
        pptConvert.PPTToJPG(PPTConvertor.SaveToType.PPT_SAVEAS_JPG);
        long e = System.currentTimeMillis();
        log.info("----PPT开始转换为Image, 成功, 共耗时 {}秒", new Object[] { new Long((e - s) / 1000) });
    }
}
