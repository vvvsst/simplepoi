package org.simplepoi.test;

import org.junit.Test;
import org.simplepoi.excel.ExcelImportUtil;
import org.simplepoi.excel.ImportParams;
import org.simplepoi.test.entity.BasicTestEntity;

import java.io.InputStream;
import java.util.List;

public class ImportTest {

    // width doesn't work from time to time

    // * For basic import test
    // read to LocalDate, Date, Decimal, Double, Integer, Long
    // read various data format, like prefix, or suffix string, . read formula.
    // read Numeric, Date format
    // read current row number
    // replace certain predefined string/characters to integer
    // drop full-null object automatically, with use of reflection util if needed
    @Test
    public void basicTest() {
        // ignore head row
        List<BasicTestEntity> resultList = null;
        try (InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("BasicTest1.xlsx")) {
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

    // Excel


    // * For one to many cases
    // ImportParams . keyIndex
    @Test
    public void oneToManyTest(){
        System.out.println("ok");
    }


    // * For exception handle test
    // read exceptional Wrong value, handle the exception elegantly


}
