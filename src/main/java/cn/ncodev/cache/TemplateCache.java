package cn.ncodev.cache;

import freemarker.cache.ByteArrayTemplateLoader;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateNotFoundException;
import freemarker.template.Version;

import java.io.*;
import java.util.Arrays;

/**
 * @author wenbing.li
 * @date 2020/6/24 15:28
 * @since 1.0
 **/
public class TemplateCache {
    /**
     * 线程模板加载参数
     */
    public static final ThreadLocal<FileLoader> LOCAL_TEMPLATE_LOADER;

    public static Configuration TEMPLATE_CONFIG;

    static {
        LOCAL_TEMPLATE_LOADER = new ThreadLocal<>();
    }

    /**
     * 获取模板的输入流
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @return 模板输入流
     * @exception  IOException  if an I/O error occurs. In particular,
     *             an <code>IOException</code> may be thrown if the
     *             output stream has been closed.
     */
    public static InputStream getTemplateStream(String template) throws IOException {
        byte[] result;
        //复杂数据,防止操作原数据
        if (LOCAL_TEMPLATE_LOADER.get() != null) {
            result = LOCAL_TEMPLATE_LOADER.get().loaderTemplate(template);
        } else {
            FileLoader fileLoader = new FileLoader();
            LOCAL_TEMPLATE_LOADER.set(fileLoader);
            result = fileLoader.loaderTemplate(template);
        }
        result = Arrays.copyOf(result, result.length);
        return new ByteArrayInputStream(result);
    }

    /**
     * 获取freemarker模板
     * @param template 模板地址
     * @param version freemarker版本
     * @return 模板
     * @throws IOException if an I/O error occurs. In particular,
     *         an <code>IOException</code> may be thrown if the
     *         output stream has been closed.
     */
    public static Template getFlTemplate(String template, Version version) throws IOException {
        if(TEMPLATE_CONFIG == null){
            TEMPLATE_CONFIG = new Configuration(version);
            TEMPLATE_CONFIG.setDefaultEncoding("UTF-8");
        }
        String templateKey = template.replace(File.separator,"_").replace("//","_").replace("/","_").replace(".","_");
        Template temp;
        try {
            temp = TEMPLATE_CONFIG.getTemplate(templateKey);
        } catch (TemplateNotFoundException e){
            FileLoader fileLoader = new FileLoader();
            byte[] result = fileLoader.loaderTemplate(template);
            ByteArrayTemplateLoader templateLoader = new ByteArrayTemplateLoader();
            templateLoader.putTemplate(templateKey,result);
            TEMPLATE_CONFIG.setTemplateLoader(templateLoader);
            temp = TEMPLATE_CONFIG.getTemplate(templateKey);
        }
        return temp;
    }
}
