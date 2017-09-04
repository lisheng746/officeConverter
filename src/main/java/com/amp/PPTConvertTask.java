package com.amp;

import com.amp.utils.MSDocConvert;

/**
 * Created by wenyong on 8/4/17.
 */
public class PPTConvertTask implements Runnable {

    private String pptPath;
    private String pdfPath;

    public PPTConvertTask(String pptPath, String pdfPath) {
        this.pptPath = pptPath;
        this.pdfPath = pdfPath;
    }

    public void run() {
        System.out.println("开始转换文件:" + pptPath);
        long startime = System.currentTimeMillis();
        MSDocConvert.convert(pptPath, pdfPath);
        long duration = System.currentTimeMillis() - startime;
        System.out.println(pptPath+"转换成功，耗时：" + (duration/1000.0) + "s");
    }
}
