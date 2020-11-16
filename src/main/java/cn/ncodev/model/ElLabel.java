package cn.ncodev.model;

/**
 * 变量标签
 * @Author wenbing.li
 * @Date 2020/4/23 9:52
 **/
public class ElLabel {
    /**
     * 文本填充开始标签，转义使用\{
     */
    public static final String START_LABEL ="${";
    /**
     * 文本填充结束标签，转义使用\{
     */
    public static final String END_LABEL ="}";

    public static final String LABEL_REGEX = "\\$\\{(.*?)\\}";
    /**
     * 填充表格使用 例如：for(item:list)，在填充列的第一个单元格使用
     */
    public static final String TRAVERSE_LABEL =":";
    /**
     * 硬回车：创建新段落，回车后不缩进
     */
    public static final String CARRIAGE_RETURN_ESCAPE = "\r\n";
    /**
     * 软回车：不创建新段落，回车后自动缩进
     */
    public static final String SOFT_CARRIAGE_RETURN_VALUES = "\n";

}
