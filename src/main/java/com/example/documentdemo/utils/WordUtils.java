package com.example.documentdemo.utils;

import com.aspose.words.*;
import com.google.common.collect.ImmutableMap;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Map;

/**
 * aspose words
 *
 * @author wuxianglong
 */
@Slf4j
public class WordUtils {
    private static final String OS_NAME_STR = "os.name";
    private static final String WINDOWS_STR = "windows";
    private static final String FORM_TEXT = "FORMTEXT";

    /**
     * linux 字体库文件目录
     * 我这个是 Centos8 下的目录
     */
    private static final String LINUX_FONTS_PATH = "/usr/share/fonts";

    public static void main(String[] args) throws Exception {
        checkLicense();
        String inPath = "";
        String outPath = "";
        docToHtml(inPath, outPath);
    }

    /**
     * word转html
     *
     * @param inPath  输入文件路径
     * @param outPath 输出文件路径
     * @throws Exception 操作异常
     */
    public static void docToHtml(String inPath, String outPath) throws Exception {
        long start = System.currentTimeMillis();
        Document doc = new Document(inPath);
        HtmlSaveOptions opts = new HtmlSaveOptions(SaveFormat.HTML);
        opts.setHtmlVersion(HtmlVersion.XHTML);
        opts.setExportImagesAsBase64(true);
        opts.setExportPageMargins(true);
        opts.setExportXhtmlTransitional(true);
        opts.setExportDocumentProperties(true);
        doc.save(outPath, opts);
        log.info("WORD转HTML成功，耗时：{}", System.currentTimeMillis() - start);
    }

    /**
     * html转word
     *
     * @param inPath  输入文件路径
     * @param outPath 输出文件路径
     * @throws Exception 操作异常
     */
    public static void htmlToWord(String inPath, String outPath) throws Exception {
        Document wordDoc = new Document(inPath);
        DocumentBuilder builder = new DocumentBuilder(wordDoc);
        for (Field field : wordDoc.getRange().getFields()) {
            if (FORM_TEXT.contains(field.getFieldCode())) {
                // 去除掉文字型窗体域
                builder.moveToField(field, true);
                builder.write(field.getResult());
                field.remove();
            }
        }
        wordDoc.save(outPath, SaveFormat.DOCX);
    }

    /**
     * html转word，并替换指定字段内容
     *
     * @param inPath  输入文件路径
     * @param outPath 输出文件路径
     * @throws Exception 操作异常
     */
    public static void htmlToWordAndReplaceField(String inPath, String outPath) throws Exception {
        Document wordDoc = new Document(inPath);
        Range range = wordDoc.getRange();
        // 把张三替换成李四，把20替换成40
        ImmutableMap<String, String> map = ImmutableMap.of("张三", "李四", "20", "40");
        for (Map.Entry<String, String> str : map.entrySet()) {
            range.replace(str.getKey(), str.getValue(), new FindReplaceOptions());
        }
        wordDoc.save(outPath, SaveFormat.DOCX);
    }

    /**
     * word转pdf
     *
     * @param inPath  输入文件路径
     * @param outPath 输出文件路径
     * @throws Exception 操作异常
     */
    public void docToPdf(String inPath, String outPath) throws Exception {
        long start = System.currentTimeMillis();
        log.info("WORD转PDF保存路径:{}", outPath);
        File file = new File(outPath);
        FileOutputStream os = new FileOutputStream(file);
        if (!System.getProperty(OS_NAME_STR).toLowerCase().startsWith(WINDOWS_STR)) {
            // linux 需要配置字体库
            log.info("【WordUtils -> docToPdf】linux字体库文件路径:{}", LINUX_FONTS_PATH);
            FontSettings.getDefaultInstance().setFontsFolder(LINUX_FONTS_PATH, false);
        }
        // 读原始文档
        Document doc = new Document(inPath);
        // 转 pdf
        doc.save(os, SaveFormat.PDF);
        os.close();
        log.info("WORD转PDF成功，耗时：{}", System.currentTimeMillis() - start);
    }

    /**
     * 校验许可文件
     */
    private static void checkLicense() {
        try {
            InputStream is = com.aspose.words.Document.class.getResourceAsStream("/com.aspose.words.lic_2999.xml");
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
