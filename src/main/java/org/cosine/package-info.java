/**
 * word2007填充，参照easypoi中的word填充进行修改，修改列表图片填充
 * https://gitee.com/lemur/easypoi
 * <!-- 需要引入 -->
 * <dependency>
 *    <groupId>org.apache.poi</groupId>
 *    <artifactId>ooxml-schemas</artifactId>
 *    <version>1.3</version>
 * </dependency>
 * <dependency>
 *     <groupId>org.apache.poi</groupId>
 *     <artifactId>poi-ooxml</artifactId>
 *     <version>3.17</version>
 * </dependency>
 *
 * 使用说明：
 * 普通文本填充：在word中使用{{}}标记要填充的参数，并在map中设置值；例如：{{name}} map.put("name","普通字段填充")
 * 文本回车换行：
 *  \r\n：硬回车：创建新段落，回车后不缩进
 *  \n ：软回车：不创建新段落，回车后自动缩进
 * 图片填充：标记方法同普通文本一样，map的值是com.gishere.utils.word.wordfill.model.WordImage类型；
 * 列表填充：如下格式
 * |{{t::list t.name|t.age|t.sex}}|
 */
package org.cosine;
