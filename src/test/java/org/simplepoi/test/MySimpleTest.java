package org.simplepoi.test;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.simplepoi.excel.ExcelExportUtil;
import org.simplepoi.excel.ExcelImportUtil;
import org.simplepoi.excel.ExportParams;
import org.simplepoi.excel.ImportParams;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.test.entity.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.management.GarbageCollectorMXBean;
import java.lang.management.ManagementFactory;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class MySimpleTest {

    @Test
    public void test1() {
        System.out.println(Integer.valueOf("1"));
    }

    @Test
    public void exportTest() {
        Teacher teach1 = new Teacher("教师1", "15.22", "222");
        Teacher teach2 = new Teacher("教师2", "111", "222");

        Student student1 = new Student("学生1", "18岁", "男");
        Student student2 = new Student("学生2", "15岁", "男");
        Student student3 = new Student("学生3", "18岁", "女");
        Student student4 = new Student("学生4", "23岁", "女");
        Student student5 = new Student("学生5", "27岁", "男");
        Student student6 = new Student("学生6", "28岁", "男");

        ArrayList<Student> students1 = new ArrayList<>();
        students1.add(student1);
        students1.add(student2);
        students1.add(student3);
        ArrayList<Student> students12 = new ArrayList<>();
        students12.add(student1);
        students12.add(student1);
        ArrayList<Student> students2 = new ArrayList<>();
        students2.add(student4);
        students2.add(student5);
        students2.add(student6);
        ArrayList<Student> students22 = new ArrayList<>();
        students22.add(student1);
        students22.add(student5);
        students22.add(student5);
        students22.add(student5);
        students22.add(student5);

        teach1.setStudentList(students1);
        teach1.setStudentList2(students12);
        teach2.setStudentList(students2);
        teach2.setStudentList2(students22);

        ArrayList<Teacher> dataList = new ArrayList<>();
        dataList.add(teach1);
        dataList.add(teach2);

        Workbook wb = ExcelExportUtil.exportExcel(new ExportParams(null, "教师信息", ExcelType.XSSF), Teacher.class, dataList);
        exportToFile(wb);
        System.out.println("ok");
    }

    @Test
    public void exportTest2() {
        Teacher2 teach1 = new Teacher2("教师1", "15.22", "222");

        Student student1 = new Student("学生1", "18岁", "男");
        Student student2 = new Student("学生2", "15岁", "男");
        Student student3 = new Student("学生3", "18岁", "女");
        Student student4 = new Student("学生4", "23岁", "女");
        Student student5 = new Student("学生5", "27岁", "男");
        Student student6 = new Student("学生6", "28岁", "男");
        Student2 student7 = new Student2("学生6", "28岁", "男");
        Student2 student8 = new Student2("学生7", "28岁", "男");

        ArrayList<Student> students1 = new ArrayList<>();
        students1.add(student1);
        students1.add(student2);
        students1.add(student3);
        ArrayList<Student> students12 = new ArrayList<>();
        students12.add(student1);
        students12.add(student1);
        ArrayList<Student> students2 = new ArrayList<>();
        students2.add(student4);
        students2.add(student5);
        students2.add(student6);
        ArrayList<Student> students22 = new ArrayList<>();
        students22.add(student1);
        students22.add(student5);
        students22.add(student5);
        students22.add(student5);
        students22.add(student5);

        ArrayList<Student2> students3 = new ArrayList<>();
        student7.setStudentList(students22);
        student8.setStudentList(students12);
        students3.add(student7);
        students3.add(student8);
        teach1.setStudentList(students1);
        teach1.setStudentList2(students12);
        teach1.setStudentList3(students3);

        ArrayList<Teacher2> dataList = new ArrayList<>();
        dataList.add(teach1);

        Workbook wb = ExcelExportUtil.exportExcel(new ExportParams(null, "教师信息", ExcelType.XSSF), Teacher2.class, dataList);
        exportToFile(wb);
        System.out.println("ok");

        List<GarbageCollectorMXBean> list = ManagementFactory.getGarbageCollectorMXBeans();
        for(GarbageCollectorMXBean bean : list) {
            System.out.println(bean.getName());
        }
    }

    public void exportToFile(Workbook wb) {
        File file = new File("C:\\Users\\Administrator\\Desktop\\教师信息.xlsx");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            wb.write(fileOutputStream);
            wb.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    // drop full-null object automatically, with use of reflection util if needed
    @Test
    public void importTest() {
        // ignore head row
        List<BasicTestEntity> resultList = null;
        try (InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("教师信息group.xlsx")) {
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


    @Test
    public void importTest2() {
        // ignore head row
        List<Teacher> resultList = null;
        try (InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("教师信息group.xlsx")) {
            resultList = ExcelImportUtil.importExcel(resourceAsStream, Teacher.class, new ImportParams());
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

    @Test
    public void importTest3() {
        // ignore head row
        List<Teacher> resultList = null;
        try  {
            resultList = ExcelImportUtil.importExcelFromDesktop( Teacher.class, "教师信息group.xlsx");
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

    @Test
    public void propTest(){

        Properties properties = new Properties();
        properties.setProperty("aaa","tt1");
        System.out.println(properties.getProperty("aaa"));
        properties.setProperty("aaa","tt2");
        System.out.println(properties.getProperty("aaa"));

        System.out.println("ok");

    }

    @Test
    public void charTest(){

        char cc1 = 'A';
        char rr1 = cc1;
        System.out.println( (char) (cc1  + 1));
        new Properties().put("aa", Character.toString(rr1));
        System.out.println("ok");

    }

}
