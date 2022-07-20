package com.example.documentdemo.utils;

import java.awt.image.BufferedImage;
import java.util.List;

/**
 * 图片合同工具
 *
 * @author wuxianglong
 */
public class MergeImage {

    /**
     * 合并任数量的图片成一张图片
     *
     * @param isHorizontal true代表水平合并，false代表垂直合并
     * @param images       待合并的图片数组
     * @return BufferedImage
     */
    public static BufferedImage mergeImage(boolean isHorizontal, List<BufferedImage> images) {
        // 生成新图片
        BufferedImage destImage = null;
        // 计算新图片的长和高
        int allWidth = 0, allHeight = 0, allWidthMax = 0, allHeightMax = 0;
        // 获取总长、总宽、最长、最宽
        for (int i = 0; i < images.size(); i++) {
            BufferedImage img = images.get(i);
            allWidth += img.getWidth();
            if (images.size() != i + 1) {
                allHeight += img.getHeight() + 2;
            } else {
                allHeight += img.getHeight();
            }
            if (img.getWidth() > allWidthMax) {
                allWidthMax = img.getWidth();
            }
            if (img.getHeight() > allHeightMax) {
                allHeightMax = img.getHeight();
            }
        }
        // 创建新图片
        if (isHorizontal) {
            destImage = new BufferedImage(allWidth, allHeightMax, BufferedImage.TYPE_INT_RGB);
        } else {
            destImage = new BufferedImage(allWidthMax, allHeight, BufferedImage.TYPE_INT_RGB);
        }
        // 合并所有子图片到新图片
        int wx = 0, wy = 0;
        for (BufferedImage img : images) {
            int w1 = img.getWidth();
            int h1 = img.getHeight();
            // 从图片中读取RGB
            int[] imageArrayOne = new int[w1 * h1];
            // 逐行扫描图像中各个像素的RGB到数组中
            imageArrayOne = img.getRGB(0, 0, w1, h1, imageArrayOne, 0, w1);
            if (isHorizontal) {
                // 水平方向合并
                // 设置上半部分或左半部分的RGB
                destImage.setRGB(wx, 0, w1, h1, imageArrayOne, 0, w1);
            } else {
                // 垂直方向合并
                // 设置上半部分或左半部分的RGB
                destImage.setRGB(0, wy, w1, h1, imageArrayOne, 0, w1);
            }
            wx += w1;
            wy += h1 + 2;
        }
        return destImage;
    }


}
