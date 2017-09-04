package com.amp;

import com.amp.openoffice.OfficeConverter;

public class OpenOfficeTest {

    public static final String pathPfx = "C:/Users/Administrator/Desktop/office_converter/doconvert";

    public static void main(String[] args) {

        String source = pathPfx +"/ppt/计算机组成原理与汇编语言程序设计第5章.ppt";
//        String target = pathPfx +"/html/计算机组成原理与汇编语言程序设计第5章.html";
        String target = pathPfx +"/pdf/计算机组成原理与汇编语言程序设计第5章.pdf";

        OfficeConverter.office2PDF(source, target);

    }

}
