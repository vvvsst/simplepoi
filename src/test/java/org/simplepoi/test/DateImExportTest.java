package org.simplepoi.test;

import org.apache.poi.ss.format.CellTextFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.simplepoi.excel.ExcelExportUtil;
import org.simplepoi.excel.ExcelImportUtil;
import org.simplepoi.excel.ExportParams;
import org.simplepoi.excel.ImportParams;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.test.entity.DateEntity;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DateImExportTest {

    @Test
    public void exportTest() {

        ArrayList<DateEntity> dataList = new ArrayList<>();
        dataList.add(new DateEntity(new Date(), LocalDate.now(), new Date(), LocalDate.now()));
        Workbook wb = ExcelExportUtil.exportExcel(new ExportParams(null, "教师信息", ExcelType.XSSF), DateEntity.class, dataList);
        exportToFile(wb);
        System.out.println("ok");

    }

    public void exportToFile(Workbook wb) {
        File file = new File("C:\\Users\\Administrator\\Desktop\\日期测试.xlsx");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            wb.write(fileOutputStream);
            wb.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    public void importFormatTest() {
        CellTextFormatter cellTextFormatter = new CellTextFormatter("\"aaa\"-@");
        String aa = cellTextFormatter.format("aa");
        System.out.println(aa);
    }

    @Test
    public void importTest() {
        File file = new File("C:\\Users\\Administrator\\Desktop\\日期测试.xlsx");
        if (!file.exists()) {
            System.out.println("文件不存在 : " + file.getAbsolutePath());
            return;
        }
        List<DateEntity> list;
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            list = ExcelImportUtil.importExcel(fileInputStream, DateEntity.class, new ImportParams());
            fileInputStream.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (list != null) {
            for (DateEntity expressImDto : list) {
                System.out.println(expressImDto.toString());
            }
        } else {
            System.out.println("获取的 列表为空");
        }
    }
}
