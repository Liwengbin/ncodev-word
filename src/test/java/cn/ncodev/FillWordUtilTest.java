package cn.ncodev;

import cn.ncodev.model.ElLabel;
import cn.ncodev.model.WordImage;
import freemarker.template.Configuration;
import freemarker.template.TemplateException;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class FillWordUtilTest {
    private static final HashMap<String, Object> map = new HashMap<>();

    @BeforeClass
    public static void beforeClass() throws Exception {
        map.put("image",new WordImage("sftheard.jpg", "/template/image/image.jpg",72,92));
        map.put("fullName","代码牛");
        map.put("sex","未知");
        map.put("political","党员");
        map.put("identityCode","53230199xxxxxxxxx");
        map.put("birthday","1999.12");
        map.put("phone","18508789615");
        map.put("hometown","云南昆明");
        map.put("toSchool","台北大学");
        map.put("school","福建师范大学");
        map.put("profession","软件工程");
        map.put("person","代码牛");
        map.put("time","2020年7月5日");
        map.put("nation","汉族");

        // 家庭信息 1条数据
        map.put("name","张三");
        map.put("age","56");
        map.put("relation","父亲");
        map.put("familyphone","18630384323");
        map.put("job","云南大学教授");
    }

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


    @Test
    public void exportWord03() {
        map.put("resume","1994.11-2014.09 云南家里蹲土地管理员（1994.11-2014.9云南XXX人民教师抬杠员，获得抬杠金奖）\n2014.09-2020.07 云南昆明 代码搬运工");
        try {
            File file = new File("D:/temporary/03赴台学生备案资料登记表.doc");
            if (file.exists() || file.createNewFile()){
                OutputStream out = new FileOutputStream(file);
                HWPFDocument doc = FillWordUtil.exportWord03("/template/03赴台学生备案资料登记表.doc",map);
                doc.write(out);
                out.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Test
    public void exportFlWord() {
        map.put("resume","1994.11-2014.09 云南家里蹲土地管理员（1994.11-2014.9云南XXX人民教师抬杠员，获得抬杠金奖）" + "<w:p/>" + "2014.09-2020.07 云南昆明 代码搬运工");
        try {
            File file = new File("D:/temporary/03赴台学生备案资料登记表(xml).doc");
            if (file.exists() || file.createNewFile()){
                OutputStream out = new FileOutputStream(file);
                FillWordUtil.exportFlWord("/template/03赴台学生备案资料登记表.xml", Configuration.VERSION_2_3_0,out,map);
                out.close();
            }
        } catch (IOException | TemplateException e) {
            e.printStackTrace();
        }
    }

    /**
     * 03简历带缩进和悬挂格式
     */
    @Test
    public void exportFlWordResume() {
        List<String> resumes = new ArrayList<>();
        resumes.add("1994.11-2014.09 云南家里蹲土地管理员（1994.11-2014.9云南XXX人民教师抬杠员，获得抬杠金奖）");
        resumes.add("2014.09-2020.07 云南昆明 代码搬运工");
        map.put("resumes",resumes);
        try {
            File file = new File("D:/temporary/03简历带缩进和悬挂格式(xml).doc");
            if (file.exists() || file.createNewFile()){
                OutputStream out = new FileOutputStream(file);
                FillWordUtil.exportFlWord("/template/03简历带缩进和悬挂格式.xml", Configuration.VERSION_2_3_0,out,map);
                out.close();
            }
        } catch (IOException | TemplateException e) {
            e.printStackTrace();
        }
    }

    /**
     * 03家庭成员退出
     */
    @Test
    public void exportFlWordFamily() {
        List<Map<String,String>> familyList = new ArrayList<>();
        Map<String,String> family = new HashMap<>();
        family.put("name","张三");
        family.put("age","56");
        family.put("relation","父亲");
        family.put("familyphone","18630384323");
        family.put("job","云南大学教授");
        familyList.add(family);
        Map<String,String> family2 = new HashMap<>();
        family2.put("name","王翠花");
        family2.put("age","30");
        family2.put("relation","目前");
        family2.put("familyphone","17639473234");
        family2.put("job","云南师范大学教授");
        familyList.add(family2);

        map.put("familys",familyList);
        try {
            File file = new File("D:/temporary/03家庭成员退出(xml).doc");
            if (file.exists() || file.createNewFile()){
                OutputStream out = new FileOutputStream(file);
                FillWordUtil.exportFlWord("/template/03家庭成员退出.xml", Configuration.VERSION_2_3_0,out,map);
                out.close();
            }
        } catch (IOException | TemplateException e) {
            e.printStackTrace();
        }
    }

}
