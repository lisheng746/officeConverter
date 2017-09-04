package com.amep.utils;

import com.amp.utils.PDFConverter;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;

/**
 * Created by wenyong on 8/3/17.
 */
public class PDFConverterTests {

    @Test
    public void convertPPTToPDF(){
        String pptPath = "/Users/wenyong/Downloads/fileConversion/ppt/万维引擎宣讲-（教仪领域 ）.ppt";
        String pdfPath = "/Users/wenyong/Downloads/fileConversion/pdf/export.html";
        long startTime = System.currentTimeMillis();
        PDFConverter.convertPPT2PDF(pptPath, pdfPath);
        long duration = System.currentTimeMillis() - startTime;

        File pdfFile = new File(pdfPath);
        Assert.assertTrue(pdfFile.exists());
        System.out.println("转换时长：" + (duration/1000.0) + "s");
    }
}
