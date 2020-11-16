# cosine-word
>简介：
 * Word填充工具(fill word tools)，支持03、07版本的Word填充。
 * word03版本填充存在换行问题，可使用xml方式替换。

> 一. xml方式填充Word
 * 需要学习fl语法 http://freemarker.foofun.cn/

> 二、使用说明
 * 普通文本填充：在word中使用{{}}标记要填充的参数，并在map中设置值；例如：{{name}} map.put("name","普通字段填充")
 * 文本回车换行：
 *  \r\n：硬回车：创建新段落，回车后不缩进
 *  \n ：软回车：不创建新段落，回车后自动缩进
 * 图片填充：标记方法同普通文本一样，map的值是org.cosine.model.WordImage类型；
 * 列表填充：如下格式
 * |{{t::list t.name|t.age|t.sex}}|
 
 >三、具体使用方法
 * 请查看测试类`org.cosine.FillWordUtilTest`
