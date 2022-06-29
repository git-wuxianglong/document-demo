package com.example.documentdemo.utils;

import com.aspose.pdf.Document;
import com.aspose.pdf.License;
import com.aspose.pdf.SaveFormat;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;

/**
 * @author wuxianglong
 */
@Slf4j
public class PdfUtils {

    public static void main(String[] args) {
        checkLicense();
        String inPath = "C:\\Users\\wuxianglong\\Desktop\\test.pdf";
        String outPath = "C:\\Users\\wuxianglong\\Desktop\\test.html";
        pdfToHtml(inPath, outPath);
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
