package org.cosine;

import org.apache.poi.hwpf.HWPFDocument;
import org.cosine.model.WordImage;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

public class FillWordUtilTest {
    private static final HashMap<String, Object> map = new HashMap<>();

    @BeforeClass
    public static void beforeClass() throws Exception {
        map.put("image",new WordImage("sftheard.jpg", "/template/image/sftheard.jpg"));
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
        map.put("resume","1994.11-2014.09 云南家里蹲土地管理员（1994.11-2014.9云南XXX人民教师抬杠员）<w:p></w:p>" +
                "2014.09-2020.07 云南昆明 代码搬运工\n");
    }

    @Test
    public void exportWord07() {

    }


    @Test
    public void exportWord03() {
        try {
            File file = new File("D:/temporary/赴台学生备案资料登记表.doc");
            if (file.exists() || file.createNewFile()){
                OutputStream out = new FileOutputStream(file);
                HWPFDocument doc = FillWordUtil.exportWord03("/template/赴台学生备案资料登记表.doc",map);
                doc.write(out);
                out.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Test
    public void exportFlWord03() {

    }
}
