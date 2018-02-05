package com.fantom.excel;

import com.fantom.excel.support.JsonUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelIOTest {

    @Test
    public void testAll() {
        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("name", "tom");
        map.put("age", 18);
        mapList.add(map);
//        File file = new File("/Users/fantom/Documents/excelTest.csv");
        File file = new File("D:\\excelTest.csv");
        try {
            ExcelIO.getWriter().generate(mapList, "mySheet")
                    .write(file);
            List<Map<String, Object>> result = ExcelIO.getReader().read(file).toMaplist();
            System.out.println(JsonUtil.obj2Json(result));
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }
}
