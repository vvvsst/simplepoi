package org.simplepoi.functest;

import org.junit.Test;
import org.simplepoi.excel.ExcelImportUtil;
import org.simplepoi.excel.ImportParams;
import org.simplepoi.test.entity.BasicTestEntity;

import java.io.InputStream;
import java.util.List;

public class ImportExport1Test {


    @Test
    public void test1() {
        // import from excel

        // ignore head row
        List<BasicTestEntity> resultList = null;
        try (InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("./functest/ImportExcel1.xlsx")) {
            resultList = ExcelImportUtil.importExcel(resourceAsStream, BasicTestEntity.class, new ImportParams());
            System.out.println("ok .");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (resultList == null) throw new RuntimeException(" file read error");

        for (BasicTestEntity basicTestEntity : resultList) {
            System.out.println(basicTestEntity);
        }

        System.out.println("ok");
    }
}
