package org.simplepoi.test.tree_exporttest;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.excel.export.ExcelExportServer;
import org.simplepoi.excel.export.ExcelSheetManipulator;
import org.simplepoi.test.entity.Teacher;
import org.simplepoi.excel.ReflectionUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ExcelManipulatorTest {
    @Test
    public void treeDataTest(){
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("test111");
        ExcelSheetManipulator manipulator = new ExcelSheetManipulator(sheet, ExcelType.XSSF);

        manipulator.insertOneMergedRow(prepareData1(1));
        manipulator.insertOneMergedRow(prepareData1(2));
        manipulator.insertOneMergedRow( prepareData2(3));

//        manipulator.insertOneMergedRow( );
        exportToFile(workbook);
        System.out.println("ok");
    }

    public void exportToFile(Workbook wb) {
        File file = new File("C:\\Users\\Administrator\\Desktop\\manipulator_test.xlsx");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            wb.write(fileOutputStream);
            wb.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private List<ExcelSheetManipulator.ObjElement>  prepareData1(int n){
        List<ExcelSheetManipulator.ObjElement> rowData = new ArrayList<>();
        ExcelSheetManipulator.ObjElement field1 = new ExcelSheetManipulator.ObjElement("tt"+n);
        ExcelSheetManipulator.ObjElement field2 = new ExcelSheetManipulator.ObjElement("tt"+n);
        rowData.add(field1);
        rowData.add(field2);
        return rowData;
    }

    private List<ExcelSheetManipulator.ObjElement>  prepareData2(int n){
        List<ExcelSheetManipulator.ObjElement> rowData = new ArrayList<>();
        ExcelSheetManipulator.ObjElement field1 = new ExcelSheetManipulator.ObjElement("tt"+n);
        ExcelSheetManipulator.ObjElement field2 = new ExcelSheetManipulator.ObjElement("tt"+n);
        rowData.add(field1);
        rowData.add(field2);

        List<ExcelSheetManipulator.ObjElement> objElements = prepareData1(31);
        List<ExcelSheetManipulator.ObjElement> objElements1 = prepareData1(32);
        ExcelSheetManipulator.ObjElement parentNode = new ExcelSheetManipulator.ObjElement("parent_node");
        parentNode.addSubObj(objElements);
        parentNode.addSubObj(objElements1);
        rowData.add(parentNode);
        return rowData;
    }


    @Test
    public void titleExportTest() throws Exception {
        List<ExcelExportServer.ExcelExportEntity> excelParams = new ArrayList<>(); // correspond to every field to be exported
        Class<?> pojoClass = Teacher.class;
        Field[] fields = ReflectionUtil.getClassFields(pojoClass);
        ExcelExportServer exportServer = new ExcelExportServer();
        ArrayList<ExcelSheetManipulator.HeaderElement> headerElements = new ArrayList<>();
        exportServer.readAllExcelFields(fields,excelParams,pojoClass,headerElements);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("test111");
        ExcelSheetManipulator manipulator = new ExcelSheetManipulator(sheet, ExcelType.XSSF);

        manipulator.createTitleAndHeaderRow(headerElements,"aa","bb");

        exportToFile(workbook);
        System.out.println("ok");
    }

}
