package org.cosine.cache;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.Version;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author wenbing.li
 * @date 2020/6/24 15:28
 * @since 1.0
 **/
public class TemplateManager {
    /**
     * 获取模板
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @return XWPFDocument
     */
    public static XWPFDocument getXWPFDocument(String template) throws IOException {
        try (InputStream inputStream = TemplateCache.getTemplateStream(template)) {
            return new XWPFDocument(inputStream);
        }
    }

    /**
     * 获取模板
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @return XWPFDocument
     */
    public static HWPFDocument getHWPFDocument(String template) throws IOException {
        try (InputStream inputStream = TemplateCache.getTemplateStream(template)) {
            return new HWPFDocument(inputStream);
        }
    }

    /**
     * 获取模板
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @param version freemarker的版本号 例如方式配置：Configuration.VERSION_2_3_0
     */
    public static Template getFlTemplate(String template, Version version) throws IOException {
        return TemplateCache.getFlTemplate(template,version);
    }
}
