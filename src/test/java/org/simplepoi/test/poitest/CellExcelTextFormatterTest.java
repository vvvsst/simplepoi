package org.simplepoi.test.poitest;

import org.apache.poi.ss.format.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.simplepoi.excel.imports.ExcelTextFormatter;

import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Iterator;

public class CellExcelTextFormatterTest {

    //  CellTextFormatter
//    String dataFormatString = cell.getCellStyle().getDataFormatString();
//    CellTextFormatter cellTextFormatter = new CellTextFormatter(dataFormatString);
//    result = cellTextFormatter.format(cell.getStringCellValue());

    @Test
    public void test1() throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        boolean isXSSFWorkbook = false;
        Workbook book = obtainBook("FormatTest.xlsx");
        if (book instanceof XSSFWorkbook) {
            isXSSFWorkbook = true;
        }

        Sheet sheetAt = book.getSheetAt(0);

        Iterator<Row> rows = sheetAt.rowIterator();
        while (rows.hasNext()) {

            Row row = rows.next();
            Cell cell = row.getCell(0);

//            Cell cell1 = row.getCell(1);
//            System.out.println(  cell.getStringCellValue()+  "   " + cell1.getStringCellValue() );
            CellType cellType = cell.getCellType();

            if (cellType == CellType.STRING || cellType == CellType.BLANK) {
                DataFormatter formatter2 = new DataFormatter();

                CellTextFormatter cellTextFormatter = new CellTextFormatter(cell.getCellStyle().getDataFormatString());
                CellFormatPart cellFormatPart = null;
                try {
                    String dataFormatString = cell.getCellStyle().getDataFormatString();
                    if (!dataFormatString.toLowerCase().contains("general"))cellFormatPart = new CellFormatPart(dataFormatString);
                } catch (Exception e) {
                    System.out.println(e.getMessage());
                }

                Method getCellFormatType = CellFormatPart.class.getDeclaredMethod("getCellFormatType");
                getCellFormatType.setAccessible(true);

                if (cellFormatPart != null) {
                    CellFormatType typeResult = (CellFormatType) getCellFormatType.invoke(cellFormatPart);
                    System.out.println(typeResult == CellFormatType.TEXT);
                    System.out.println(typeResult);
                } else {
                    System.out.println(false);
                }

                String format = cellTextFormatter.format(cell.getStringCellValue());
                System.out.println(cell.getStringCellValue() + " | "
                        + cell.getCellStyle().getDataFormatString() + " | "
                        + format);
            }

            if (cellType == CellType.NUMERIC) {
                DataFormatter formatter = new DataFormatter();
//                Format format1 = formatter.createFormat(cell);
//                String cellValue = formatter.formatCellValue(cell);
//                System.out.println(cellValue);
//                System.out.println(format1.format(cell.getNumericCellValue()));
//                CellNumberFormatter cellNumberFormatter = new CellNumberFormatter(cell.getCellStyle().getDataFormatString());
//                CellFormatPart cellFormatPart = new CellFormatPart(cell.getCellStyle().getDataFormatString());
//                String format = cellTextFormatter.format(cell.getStringCellValue());
//                System.out.println(cell.getNumericCellValue() + " | "
//                        + cell.getCellStyle().getDataFormatString() + " | " +
//                        cellFormatPart.apply(cell.getNumericCellValue()).text);

                System.out.println(cell.getNumericCellValue() + " | "
                        + cell.getCellStyle().getDataFormatString() + " | "
                        + formatter.formatCellValue(cell));

            }
        }
        System.out.println();
        System.out.println("ok");
    }


    @Test
    public void test2() throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        ExcelTextFormatter excelTextFormatter = ExcelTextFormatter.getInstance();
        boolean isXSSFWorkbook = false;
        Workbook book = obtainBook("FormatTest.xlsx");
        if (book instanceof XSSFWorkbook) {
            isXSSFWorkbook = true;
        }

        Sheet sheetAt = book.getSheetAt(0);

        Iterator<Row> rows = sheetAt.rowIterator();
        while (rows.hasNext()) {
            Row row = rows.next();
            Cell cell = row.getCell(0);
            CellType cellType = cell.getCellType();
            if (cellType == CellType.STRING || cellType == CellType.BLANK) {

                String format = excelTextFormatter.format(cell.getStringCellValue(),cell.getCellStyle().getDataFormatString());
                System.out.println(cell.getStringCellValue() + " | "
                        + cell.getCellStyle().getDataFormatString() + " | "
                        + format);
            }

            if (cellType == CellType.NUMERIC) {
                DataFormatter formatter = new DataFormatter();
                System.out.println(cell.getNumericCellValue() + " | "
                        + cell.getCellStyle().getDataFormatString() + " | "
                        + formatter.formatCellValue(cell));

            }
        }
        System.out.println();
        System.out.println("ok");
    }
    private Workbook obtainBook(String filename) {
        try (InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream(filename)) {
            return WorkbookFactory.create(resourceAsStream);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

}
