package org.simplepoi.test;

import org.junit.Test;
import org.simplepoi.excel.ReflectionUtil;
import org.simplepoi.excel.imports.ExcelImportServer;
import org.simplepoi.test.entity.Teacher;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.simplepoi.excel.imports.ExcelImportServer.getAllExcelField;

public class ExcelImportEntityTest {

    @Test
    public void test1() throws Exception {
        Map<String, ExcelImportServer.ExcelImportEntity> excelParams = new HashMap<>(); // correspond to every field to be exported
        List<ExcelImportServer.ExcelCollectionParams> excelCollection = new ArrayList<>();
//        ImportServerSupport.getAllExcelField()
//        Class<?> pojoClass = BasicTestEntity.class;
        Class<?> pojoClass = Teacher.class;
        Field[] fields = ReflectionUtil.getClassFields(pojoClass);
        // excelParams include all the information on the corresponding field
        getAllExcelField(fields, excelParams, excelCollection, pojoClass);

        System.out.println("ok");
    }

}
