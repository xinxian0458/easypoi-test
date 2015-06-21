package org.jeecgframework.poi.test.excel.read;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.test.entity.CustomerEntity;
import org.junit.Test;

public class CustomerEntityTest {

    @Test
    public void test() {
        /*try {
            ImportParams params = new ImportParams();
            long start = new Date().getTime();
            List<CustomerEntity> list = ExcelImportUtil.importExcel(new FileInputStream(
                new File("d:/tt.xlsx")), CustomerEntity.class, params);
            System.out.println(list.size() + "-----" + (new Date().getTime() - start));
        } catch (Exception e) {
            e.printStackTrace();
        }*/

        String t = "5.0E-7";
        long start = System.nanoTime();
        for (int i = 0; i < 100000; i++) {
            new BigDecimal(t);
        }
        System.out.println(System.nanoTime() - start);
        double d = 5.0E-7;
        long start2 = System.nanoTime();
        for (int i = 0; i < 100000; i++) {
            new BigDecimal(d);
        }
        System.out.println(System.nanoTime() - start2);
    }

}
