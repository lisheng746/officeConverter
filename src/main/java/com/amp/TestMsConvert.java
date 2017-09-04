package com.amp;

/**
 * Created by wenyong on 8/4/17.
 */
public class TestMsConvert {

    public static final String pathPfx = "C:/Users/Administrator/Desktop/office_converter/doconvert";

    public static void main(String[] args) {
//        System.out.println(System.getProperty("java.library.path"));
//        String[] pptPath = new String[]{pathPfx +"/ppt/《肌肉》PPT.ppt", pathPfx +"/ppt/01.静力学公理.pptx", pathPfx +"/ppt/计算机组成原理与汇编语言程序设计第5章.ppt"};
//        String[] pdfPath = new String[]{pathPfx +"/pdf/《肌肉》PPT..pdf", pathPfx +"/pdf/静力学公理.pdf", pathPfx +"/pdf/计算机组成原理与汇编语言程序设计第5章.pdf"};
        String[] pptPath = new String[]{pathPfx +"/ppt/01.静力学公理.pptx"};
        String[] pdfPath = new String[]{pathPfx +"/pdf/01.静力学公理.pdf"};
        int i = 0;
        for(;i<pptPath.length;i++){
            Thread thread = new Thread(new PPTConvertTask(pptPath[i], pdfPath[i]));
            thread.start();
            System.out.println("转换任务"+i+"启动");
        }
    }
}
