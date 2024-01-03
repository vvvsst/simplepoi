package org.simplepoi.test.treeimporttest;

import org.junit.Test;
import org.simplepoi.excel.ExcelImportUtil;
//import org.simplepoi.test.treeimporttest.entity.Teacher;
import org.simplepoi.excel.imports.ExcelImportServer;
import org.simplepoi.test.entity.Teacher;
import org.simplepoi.excel.ReflectionUtil;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.simplepoi.excel.imports.ExcelImportServer.getAllExcelField;

public class MultiLevelTreeImportTest {

    @Test
    public void test1() throws Exception {

        Map<String, ExcelImportServer.ExcelImportEntity> excelParams = new HashMap<>(); // correspond to every field to be exported
        List<ExcelImportServer.ExcelCollectionParams> excelCollection = new ArrayList<>();
        Class<?> pojoClass = Teacher.class;
        Field[] fields = ReflectionUtil.getClassFields(pojoClass);
        getAllExcelField(fields, excelParams, excelCollection, pojoClass);
        System.out.println("ok");
    }

    @Test
    public void importTest1() {
        // ignore head row
        List<Teacher> resultList = null;
        try {
            resultList = ExcelImportUtil.importExcelFromDesktop(Teacher.class, "教师信息group.xlsx");
            System.out.println("ok .");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (resultList == null) throw new RuntimeException(" file read error");

        for (Teacher basicTestEntity : resultList) {
            System.out.println(basicTestEntity);
        }

        System.out.println("ok");
    }
}
