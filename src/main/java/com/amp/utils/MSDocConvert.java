package com.amp.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * Created by wenyong on 8/4/17.
 */
public class MSDocConvert {
    private static final int ppFormatPDF = 32;

    public static void convert(String pptFile, String pdf) {
        //ComThread.InitMTA(true);
        ComThread.InitSTA();
        String programId = "PowerPoint.Application";
        ActiveXComponent activeXComponent = new ActiveXComponent(programId);
        //activeXComponent.setProperty("Visible", new Variant(false));
        Dispatch ppts = activeXComponent.getProperty("Presentations").toDispatch();
        boolean readOnly = true;
        boolean untitiled = true;
        boolean withWindow = false;
        Dispatch pptDispatch = Dispatch.call(ppts, "Open", pptFile, readOnly, untitiled, withWindow).toDispatch();
        Dispatch.call(pptDispatch, "SaveAs", pdf, ppFormatPDF);
        Dispatch.call(pptDispatch, "Close");
        activeXComponent.invoke("Quit", new Variant[]{});
        ComThread.Release();
        ComThread.quitMainSTA();
    }
}
