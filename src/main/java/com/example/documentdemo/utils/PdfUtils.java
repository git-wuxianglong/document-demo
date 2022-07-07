package com.example.documentdemo.utils;

import com.aspose.pdf.*;
import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;


import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author wuxianglong
 */
@Slf4j
public class PdfUtils {

    public static void main(String[] args) {
        checkLicense();
//        String inPath = "C:\\Users\\xxxx\\Desktop\\test.pdf";
//        String outPath = "C:\\Users\\xxxx\\Desktop\\test.html";
//        pdfToHtml(inPath, outPath);

        ArrayList<String> strings = Lists.newArrayList("C:\\Users\\xxxx\\Desktop\\11\\001o.png",
                "C:\\Users\\xxxx\\Desktop\\11\\002o.png",
                "C:\\Users\\xxxx\\Desktop\\11\\003o.png");
        String pdfPath = "C:\\Users\\xxxx\\Desktop\\imagePdf.pdf";
        imageToPdf(strings, pdfPath);
    }

    /**
     * PDF转HTML
     *
     * @param inPath  输入文件路径
     * @param outPath 输出文件路径
     */
    public static void pdfToHtml(String inPath, String outPath) {
        long start = System.currentTimeMillis();
        try (Document document = new Document(inPath)) {
            document.save(outPath, SaveFormat.Html);
        }
        log.info("PDF转HTML，耗时：{}", System.currentTimeMillis() - start);
    }

    public static void imageToPdf(List<String> imagePathList, String outPath) {
        long start = System.currentTimeMillis();
        try (Document doc = new Document()) {
            for (String imagePath : imagePathList) {
                Page page = doc.getPages().add();
                try (FileInputStream fs = new FileInputStream(imagePath)) {
                    page.getPageInfo().getMargin().setBottom(0);
                    page.getPageInfo().getMargin().setTop(0);
                    page.getPageInfo().getMargin().setLeft(0);
                    page.getPageInfo().getMargin().setRight(0);
                    page.setCropBox(new Rectangle(0, 0, 800, 1000));
                    Image image = new Image();
                    page.getParagraphs().add(image);
                    image.setImageStream(fs);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            doc.save(outPath);
        }
        log.info("多图片转PDF，耗时：{}", System.currentTimeMillis() - start);
    }

    /**
     * 验证License
     */
    private static void checkLicense() {
        try {
            // 注意包名
            InputStream is = com.aspose.pdf.Document.class.getResourceAsStream("/com.aspose.pdf.lic.xml");
            if (is == null) {
                return;
            }
            License asposeLicense = new License();
            asposeLicense.setLicense(is);
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
