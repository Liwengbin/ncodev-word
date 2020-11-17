> 一、简介
* 功能需求通过填充生成.doc和.docx格式的文档，由于使用easypoi内容太多（冗余），并且生成word的图片，文本换行不易处理。我将word部分代码抽离出来并对代码进行处理逻辑调整。生成单独、简单、易用的工具包`ncodev-word`

> 二、解决的问题，使用场景
* 图片填充问题，文本硬回车和软回车问题。在生成个人信息文档，简历文档时可以很轻松的使用工具来填充生成

> 三、使用说明
* 普通文本填充：在word中使用{{}}标记要填充的参数，并在map中设置值；例如：{{name}} map.put("name","普通字段填充")
* 文本回车换行：
*  rn：硬回车：创建新段落，回车后不缩进
*  n ：软回车：不创建新段落，回车后自动缩进
* 图片填充：标记方法同普通文本一样，map的值是`cn.ncodev.model.WordImage`类型；
* 列表填充：如下格式
* |{{t::list t.name|t.age|t.sex}}|
> 具体如何使用可以查看测试类：`cn.ncodev.FillWordUtilTest`

> 使用示例
```java
/**
 * 07简历带缩进和悬挂格式
 */
@Test
public void exportWord07() {
    map.put("resume","1994.11-2014.09 云南家里蹲土地管理员（1994.11-2014.9云南XXX人民教师抬杠员，获得抬杠金奖）" + ElLabel.CARRIAGE_RETURN_ESCAPE + "2014.09-2020.07 云南昆明 代码搬运工");
    try {
        File file = new File("D:/temporary/07赴台学生备案资料登记表.docx");
        if (file.exists() || file.createNewFile()){
            OutputStream out = new FileOutputStream(file);
            XWPFDocument doc = FillWordUtil.exportWord07("/template/07赴台学生备案资料登记表.docx",map);
            doc.write(out);
            out.close();
        }
    } catch (IOException e) {
        e.printStackTrace();
    }
}
```

> 安装
* GitHub:[https://github.com/Liwengbin/ncodev-word](https://github.com/Liwengbin/ncodev-word)
```xml
<dependency>
    <groupId>cn.ncodev.fill</groupId>
    <artifactId>ncodev-word</artifactId>
    <version>1.0-SNAPSHOT</version>
</dependency>
```
