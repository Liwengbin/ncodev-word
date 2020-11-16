package cn.ncodev.word;

import cn.ncodev.cache.TemplateManager;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.Version;
import cn.ncodev.cache.FileLoader;
import cn.ncodev.model.WordImage;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.util.Map;

/**
 * <w:p/> 创建段落-类似硬回车但字体样式
 * <w:br/> 软回车
 */
public class FillWord {
    /**
     * 使用freemarker模板填充word
     * @param templatePath 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @param version freemarker的版本号 例如方式配置：Configuration.VERSION_2_3_0
     * @param map 填充参数
     * @return 模板
     */
    public void fillFlWord(String templatePath, Version version, OutputStream outputStream, Map<String, Object> map) throws IOException, TemplateException {
        try (Writer out = new BufferedWriter(new OutputStreamWriter(outputStream))) {
            Template template = TemplateManager.getFlTemplate(templatePath, version);
            // 将map中的图片转为base64格式
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                if(entry.getValue() != null && entry.getValue() instanceof WordImage){
                    WordImage image = (WordImage) entry.getValue();
                    if(image.getImageBytes() == null){
                        image.setImageBytes(FileLoader.loaderFile(image.getImagePath()));
                    }
                    entry.setValue(new BASE64Encoder().encode(image.getImageBytes()));
                }
            }
            template.process(map, out);
        }
    }
}
