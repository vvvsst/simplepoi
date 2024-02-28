package org.simplepoi.functest;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.simplepoi.excel.ExcelExportUtil;
import org.simplepoi.excel.ExcelImportUtil;
import org.simplepoi.excel.ExportParams;
import org.simplepoi.excel.ImportParams;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.functest.entity.Grade;
import org.simplepoi.functest.entity.Student;
import org.simplepoi.functest.entity.Teacher;
import org.simplepoi.test.entity.BasicTestEntity;
import org.simplepoi.test.entity.WidthTestEntity1;
import org.simplepoi.test.entity.WidthTestEntity2;
import org.simplepoi.test.entity.WidthTestEntity3;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class ImportExport1Test {

    private void exportToFile(ArrayList<Map<String, Object>> listMap, String fileName) {
        Workbook workbook = ExcelExportUtil.exportExcel(listMap, ExcelType.XSSF);
        //InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("./functest/ImportExcel1.xlsx"))
        File file = new File("C:\\Users\\Administrator\\Desktop\\" + fileName + new SimpleDateFormat("yyyy-MM-dd").format(new Date()) + ".xlsx");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public <T> void exportToFile(List<T> dataList , Class<T> dataClass,String fileName) {
        ArrayList<Map<String, Object>> listMap = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("title", new ExportParams(null, "学生教师", ExcelType.XSSF));
        map.put("data", dataList);
        map.put("entity", dataClass);
        listMap.add(map);
        exportToFile(listMap, fileName);  // 文件导出
    }


    private List<Teacher> createTeacherListData(){
        Grade grade = new Grade("语文");
        Grade grade1 = new Grade("数学");
        Grade grade2 = new Grade("化学");
        Grade grade3 = new Grade("生物");

        Student student = new Student("张1");
        student.addGrade(grade).addGrade(grade1);
        Student student1 = new Student("张2");
        student1.addGrade(grade1).addGrade(grade3);
        Student student2 = new Student("张3");
        student2.addGrade(grade2).addGrade(grade3);

        Student student3 = new Student("李1");
        student3.addGrade(grade).addGrade(grade2);
        Student student4 = new Student("李2");
        student4.addGrade(grade2).addGrade(grade3);
        Student student5 = new Student("李3");
        student5.addGrade(grade2).addGrade(grade3);

        Teacher teacher1 = new Teacher("老师1");
        teacher1.addStudent1(student1).addStudent1(student2);
        teacher1.addStudent2(student3).addStudent1(student4);
        Teacher teacher2 = new Teacher("老师2");
        teacher2.addStudent1(student2).addStudent1(student3);
        teacher2.addStudent2(student4).addStudent2(student5);

        List<Teacher> teacherList = new ArrayList<>();
        teacherList.add(teacher1);
        teacherList.add(teacher2);
        return teacherList;
    }


    @Test
    public void testExport() {
        List<Teacher> teacherListData = createTeacherListData();
        ArrayList<Map<String, Object>> listMap = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("title", new ExportParams(null, "学生教师", ExcelType.XSSF));
        map.put("data", teacherListData);
        map.put("entity", Teacher.class);
        listMap.add(map);

        exportToFile(listMap, "学生教师");  // 文件导出

        System.out.println("ok");
    }

    @Test
    public void testImport() {
        List<Teacher> resultList = null;
        try (InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("./functest/学生教师.xlsx")) {
            resultList = ExcelImportUtil.importExcel(resourceAsStream, Teacher.class, new ImportParams(4,19));
            exportToFile(resultList, Teacher.class,"读入结果导出");
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
