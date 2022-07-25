package com.example.documentdemo.utils;


import com.aspose.words.*;
import com.google.common.collect.ImmutableMap;
import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
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
        String inPath = "C:\\Users\\username\\Desktop\\test.docx";
        String outPath = "C:\\Users\\username\\Desktop\\test.html";
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
     * word转pdf
     *
     * @param inPath  输入文件路径
     * @param outPath 输出文件路径
     * @throws Exception 操作异常
     */
    public static void docToPdf(String inPath, String outPath) throws Exception {
        long start = System.currentTimeMillis();
        log.info("WORD转PDF保存路径:{}", outPath);
        FileOutputStream os = getFileOutputStream(outPath);
        // 读原始文档
        Document doc = new Document(inPath);
        // 转 pdf
        doc.save(os, SaveFormat.PDF);
        os.close();
        log.info("WORD转PDF成功，耗时：{}", System.currentTimeMillis() - start);
    }

    /**
     * word转pdf
     *
     * @param inputStream 文件输入流
     * @param outPath     输出文件路径
     * @throws Exception 操作异常
     */
    public static void docToPdf(InputStream inputStream, String outPath) throws Exception {
        FileOutputStream os = getFileOutputStream(outPath);
        Document doc = new Document(inputStream);
        doc.save(os, SaveFormat.PDF);
        os.close();
    }

    /**
     * word转换为图片，每页一张图片
     *
     * @param inPath word文件路径
     * @throws Exception 操作异常
     */
    public static void docToImage(String inPath) throws Exception {
        InputStream inputStream = Files.newInputStream(Paths.get(inPath));
        File file = new File(inPath);
        String name = file.getName();
        String fileName = name.substring(0, name.lastIndexOf("."));
        String parent = file.getParent();
        log.info("parent:{}", parent);
        boolean mkdir = new File(parent + "/" + fileName).mkdir();
        log.info("mkdir:{}", mkdir);
        List<BufferedImage> bufferedImages = wordToImg(inputStream);
        for (int i = 0; i < bufferedImages.size(); i++) {
            ImageIO.write(bufferedImages.get(i), "png", new File(parent + "/" + fileName + "/" + "第" + i + "页" + fileName + ".png"));
        }
        inputStream.close();
    }

    /**
     * word转换为图片，合并为一张图片
     *
     * @param inPath word文件路径
     * @throws Exception 操作异常
     */
    public static void docToOneImage(String inPath) throws Exception {
        InputStream inputStream = Files.newInputStream(Paths.get(inPath));
        File file = new File(inPath);
        String name = file.getName();
        String fileName = name.substring(0, name.lastIndexOf("."));
        String parent = file.getParent();
        List<BufferedImage> bufferedImages = wordToImg(inputStream);
        // 合并为一张图片
        BufferedImage image = MergeImage.mergeImage(false, bufferedImages);
        ImageIO.write(image, "png", new File(parent + "/" + fileName + ".png"));
        inputStream.close();
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
            if (field.getFieldCode().contains(FORM_TEXT)) {
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
     * word转pdf，linux下设置字体库文件路径，并返回FileOutputStream
     *
     * @param outPath pdf输出路径
     * @return pdf输出路径 -> FileOutputStream
     * @throws FileNotFoundException FileNotFoundException
     */
    private static FileOutputStream getFileOutputStream(String outPath) throws FileNotFoundException {
        if (!System.getProperty(OS_NAME_STR).toLowerCase().startsWith(WINDOWS_STR)) {
            // linux 需要配置字体库
            log.info("【WordUtils -> docToPdf】linux字体库文件路径:{}", LINUX_FONTS_PATH);
            FontSettings.getDefaultInstance().setFontsFolder(LINUX_FONTS_PATH, false);
        }
        return new FileOutputStream(outPath);
    }

    /**
     * word转图片
     *
     * @param inputStream word input stream
     * @return BufferedImage list
     * @throws Exception exception
     */
    private static List<BufferedImage> wordToImg(InputStream inputStream) throws Exception {
        Document doc = new Document(inputStream);
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        options.setPrettyFormat(true);
        options.setUseAntiAliasing(true);
        options.setUseHighQualityRendering(true);
        int pageCount = doc.getPageCount();
        List<BufferedImage> imageList = new ArrayList<>();
        for (int i = 0; i < pageCount; i++) {
            OutputStream output = new ByteArrayOutputStream();
            options.setPageIndex(i);
            doc.save(output, options);
            ImageInputStream imageInputStream = ImageIO.createImageInputStream(parse(output));
            imageList.add(ImageIO.read(imageInputStream));
        }
        return imageList;
    }

    /**
     * outputStream转inputStream
     *
     * @param out OutputStream
     * @return inputStream
     */
    private static ByteArrayInputStream parse(OutputStream out) {
        return new ByteArrayInputStream(((ByteArrayOutputStream) out).toByteArray());
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
