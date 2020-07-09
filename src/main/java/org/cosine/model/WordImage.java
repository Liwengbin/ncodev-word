package org.cosine.model;

import org.apache.poi.xwpf.usermodel.Document;

/**
 * word 填充时的图片实体
 * @author wenbing.li
 * @date 2020/6/25 11:23
 */
public class WordImage {
    /**
     * 图片宽度
     */
    private int width;
    /**
     * 图片高度
     */
    private int height;
    /**
     * 图片名称，图片名称需要带后缀，如: image.jpeg
     */
    private String imageName;
    /**
     * 图片网络地址 或 （先对路径 | 绝对路径）
     */
    private String imagePath;
    /**
     * 图片字节码
     */
    private byte[] imageBytes;

    /**
     * 图片类型默认 2007 <code>Document.PICTURE_TYPE_JPEG</code>
     * @see org.apache.poi.xwpf.usermodel.Document
     */
    private int imageSuffix = Document.PICTURE_TYPE_JPEG;

    public WordImage(String imageName, String imagePath) {
        this.imageName = imageName;
        this.imagePath = imagePath;
    }

    public WordImage(String imageName, String imagePath,int width, int height) {
        this.width = width;
        this.height = height;
        this.imageName = imageName;
        this.imagePath = imagePath;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public String getImageName() {
        return imageName;
    }

    public void setImageName(String imageName) {
        this.imageName = imageName;
    }

    public String getImagePath() {
        return imagePath;
    }

    public void setImagePath(String imagePath) {
        this.imagePath = imagePath;
    }

    public byte[] getImageBytes() {
        return imageBytes;
    }

    public void setImageBytes(byte[] imageBytes) {
        this.imageBytes = imageBytes;
    }

    public int getImageSuffix() {
        return imageSuffix;
    }

    public void setImageSuffix(int imageSuffix) {
        this.imageSuffix = imageSuffix;
    }
}
