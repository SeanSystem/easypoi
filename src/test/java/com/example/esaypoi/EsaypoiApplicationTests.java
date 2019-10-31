package com.example.esaypoi;

import cn.afterturn.easypoi.word.WordExportUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import java.io.FileOutputStream;
import java.util.*;

@SpringBootTest
class EsaypoiApplicationTests {

    @Test
    void contextLoads() {
        Map<String, Object> map = new HashMap<String, Object>();
        List<Map<String,Object>> list = new ArrayList<>();

        for(int i=0;i<4;i++){
            Map<String,Object> listMap = new HashMap<String,Object>();
            listMap.put("id", i+1+"");
            listMap.put("xzq", "3302");
            listMap.put("xzqmc", "201");
            list.add(listMap);
        }
        map.put("maplist",list);
        try {
            XWPFDocument doc = WordExportUtil.exportWord07(
                    "WEB-INF/doc/word/test.docx", map);
            FileOutputStream fos = new FileOutputStream("F:/excel/simple.docx");
            doc.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
