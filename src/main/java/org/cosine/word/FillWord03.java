package org.cosine.word;

import freemarker.cache.ByteArrayTemplateLoader;
import freemarker.cache.TemplateLoader;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.Version;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.cosine.cache.TemplateManager;
import org.cosine.model.ElLabel;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * word.doc 2003版本word文档填充
 * @author wenbing.li
 * @date 2020/6/25 13:23
 */
public class FillWord03 {
    /**
     * 填充word
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @param map 填充参数
     * @return 填充后的文档
     * @exception  IOException  if an I/O error occurs. In particular,
     *             an <code>IOException</code> may be thrown if the
     *             output stream has been closed.
     */
    public HWPFDocument fillWord(String template, Map<String, Object> map) throws IOException {
        HWPFDocument hwpfDocument = TemplateManager.getHWPFDocument(template);
        this.fillWord(hwpfDocument, map);
        return hwpfDocument;
    }

    private void fillWord(HWPFDocument hwpfDocument, Map<String, Object> map) {
        if(hwpfDocument == null){
            throw new NullPointerException("HWPFDocument is null");
        }
        // 得到文档的读取范围
        Range range = hwpfDocument.getRange();
        List<String> params = parseThisTextKey(range.text());
        for (String key : params) {
            String text = "";
            if(map.containsKey(key)){
                Object value = map.get(key);
                text = String.valueOf(value);
            }
            range.replaceText(ElLabel.START_LABEL + key + ElLabel.END_LABEL,text);
        }
    }

    /**
     * 获取填充的参数
     * @param currentText 需要填充的文本
     */
    private static List<String> parseThisTextKey(String currentText) {
        Pattern pattern = Pattern.compile(ElLabel.LABEL_REGEX);
        Matcher matcher = pattern.matcher(currentText);
        List<String> params = new ArrayList<>();
        int i = 1;
        while (matcher.find()) {
            params.add(matcher.group(i));
        }
        return params;
    }

    /*word03-freemarker*/

    /**
     * @param templatePath 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @param version freemarker的版本号 例如方式配置：Configuration.VERSION_2_3_0
     * @param map 填充参数
     * @return 模板
     */
    public void fillFlWord(String templatePath, Version version, OutputStream outputStream, Map<String, Object> map) throws IOException, TemplateException {
        try (Writer out = new BufferedWriter(new OutputStreamWriter(outputStream))) {
            Template template = TemplateManager.getFlTemplate(templatePath, version);
            template.process(map, out);
        }
    }
}

