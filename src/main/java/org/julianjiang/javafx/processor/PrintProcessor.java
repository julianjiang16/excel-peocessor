package org.julianjiang.javafx.processor;

import javax.print.*;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Copies;
import javax.print.attribute.standard.JobName;
import javax.print.attribute.standard.MediaSizeName;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

public class PrintProcessor {

    public static void print(File file) {
        try {
            // 获取打印服务
            PrintService[] printServices = PrintServiceLookup.lookupPrintServices(null, null);
            if (printServices.length > 0) {
                // 用户选择打印机
                // 创建打印请求属性集
                PrintRequestAttributeSet attributes = new HashPrintRequestAttributeSet();
                attributes.add(new Copies(1)); // 设置打印份数
                attributes.add(MediaSizeName.ISO_A4); // 设置纸张大小
                attributes.add(new JobName("Print Job", null)); // 设置打印作业名称
                PrintService selectedPrintService = ServiceUI.printDialog(null, 200, 200, printServices, null, null, attributes);

                if (selectedPrintService != null) {
                    // 创建打印任务
                    DocPrintJob printJob = selectedPrintService.createPrintJob();
                    FileInputStream fis = new FileInputStream(file);
                    Doc doc = new SimpleDoc(fis, DocFlavor.INPUT_STREAM.AUTOSENSE, null);
                    printJob.print(doc, null);
                    System.out.println("文件已提交打印。");
                } else {
                    System.out.println("用户取消选择打印机。");
                }
            } else {
                System.out.println("未找到可用的打印机。");
            }
        } catch (FileNotFoundException e) {
            System.out.println("找不到要打印的文件。");
            e.printStackTrace();
        } catch (PrintException e) {
            System.out.println("无法打印文件。");
            e.printStackTrace();
        }
    }
}
