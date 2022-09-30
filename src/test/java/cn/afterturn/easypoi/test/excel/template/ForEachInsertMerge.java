package cn.afterturn.easypoi.test.excel.template;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.hutool.json.JSONUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * 插入合并这块
 * @author jueyue on 20-4-9.
 */
public class ForEachInsertMerge {

    @Test
    public void sendGoods() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "doc/foreachInsertMerge.goods_send.xlsx");
        Map<String, Object>       map     = new HashMap<>();
        map.put("nowTime",new Date());
        map.put("unitName","悟耘信息");
        map.put("order",new Date().getTime());
        List<Map<String, Object>> mapList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> testMap = new HashMap<>();
            testMap.put("name", "小明" + i);
            testMap.put("nums",i);
            testMap.put("type","食品");
            testMap.put("remark","甜食");
            mapList.add(testMap);
        }
        map.put("list", mapList);
        //本来导出是专业那个
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        File     savefile = new File("D:/home/excel/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/home/excel/ForEachInsertMerge.sendGoods.xlsx");
        workbook.write(fos);
        fos.close();
    }

    @Test
    public void jessieTest() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "doc/jessie_test.xlsx");
        Map<String, Object>       map     = new HashMap<>();
        List<Map<String, Object>> rowList = new ArrayList<>();
        for (int i=0; i<2; i++) {
            Map<String, Object> testMap = new HashMap<>();
            testMap.put("id", i);
            testMap.put("name", "jessie" + i);
            List<Map<String, Object>> mapList = new ArrayList<>();
            for (int j = 0; j < 2; j++) {
                Map<String, Object> dataMap = new HashMap<>();
                dataMap.put("year", 2022 + j);
                dataMap.put("value", 200 + i + j);
                dataMap.put("score", 100 + i + j);
                mapList.add(dataMap);
            }
            testMap.put("data", mapList);
            rowList.add(testMap);
        }
//        List<Map<String, Object>> colList = new ArrayList<>();
//        for (int i=0; i<2; i++) {
//            Map<String, Object> colMap = new HashMap<>();
//            colMap.put("year", 2022 + i);
//            colMap.put("field", "n:t.data.year==" + (2022 + i) + "?n:t.data.value: 0");
//            colList.add(colMap);
//        }
        map.put("rows", rowList);
        this.rebuildMap(map);
//        map.put("cols", colList);
        System.out.println(JSONUtil.parse(map).toStringPretty());
        //本来导出是专业那个
        params.setColForEach(true);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        File     savefile = new File("D:/home/excel/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/home/excel/Jessie_test_result.xlsx");
        workbook.write(fos);
        fos.close();
    }

    private void rebuildMap(Map<String, Object> originMap) {
        Map<String, Object> colsMap = new HashMap<>();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) originMap.get("rows");
        for(Map<String, Object> row: rows) {
            List<Map<String, Object>> datas = (List<Map<String, Object>>) row.get("data");
            for(Map<String, Object> data: datas) {
                String year = "" + data.get("year");
                row.put("t_" + year + "_value", data.get("value"));
                row.put("t_" + year + "_score", data.get("score"));
                if (!colsMap.containsKey(year)) {
                    colsMap.put(year, "t.name == 'jessie0' ? t.t_" + year + "_value" + " : t.t_" + year + "_score");
//                    colsMap.put(year, "t.id == '0' ? '男' : '女'");
                }
            }
        }
        List<Map<String, Object>> cols = new ArrayList<>();
        for (String year: colsMap.keySet()) {
            Map<String, Object> col = new HashMap<>();
            col.put("year", year);
            col.put("field", colsMap.get(year));
            cols.add(col);
        }
        originMap.put("cols", cols);
    }
}
