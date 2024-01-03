package org.simplepoi.excel.export;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.excel.constant.PoiBaseConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Properties;

import static org.simplepoi.excel.constant.PoiBaseConstants.*;
import static org.simplepoi.excel.export.ImageCellCreateSupport.createImageCell;
import static org.simplepoi.excel.export.ImageCellCreateSupport.setImageCellType;

public class ExcelSheetManipulator {
    private int finishedLine = 0;
    private Sheet sheet;
    private ExcelType type = ExcelType.HSSF;
    private List<Row> rows = new ArrayList<>();
    private List<List<ObjElement>> rowDatas = new ArrayList<>();
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelSheetManipulator.class);
    private CellStyle cellStyle;
    private Drawing drawing;
    private Properties properties = new Properties();

    private GenericTokenParser tokenParser = new GenericTokenParser("${", "}", properties);

    public Drawing getDrawaing() {
        if (drawing == null)
            return sheet.createDrawingPatriarch();
        else return drawing;
    }

    public int getFinishedLine() {
        return finishedLine;
    }

    public ExcelSheetManipulator(Sheet sheet) {
        this.sheet = sheet;
        initStyle();
    }

    public ExcelSheetManipulator(Sheet sheet, ExcelType type) {
        this.sheet = sheet;
        if (type != null) this.type = type;
        setImageCellType(this.type);
        initStyle();
    }

    private void initStyle() {
        cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
        sheet.setDefaultRowHeight((short) (20 * 20));
    }

    // width of every column
    // header row
    @Deprecated // done by createHeaderRow
    public ExcelSheetManipulator columnWidth(List<Integer> widthForColumn) {
        return this;
    }

//    public ExcelSheetManipulator columnHidden(List<Integer> hiddenColumns) {
//        return this;
//    }

    public ExcelSheetManipulator createTitleAndHeaderRow(List<HeaderElement> rowData, String title, String secondTitle) {
        // all type should be  String
        // title = null or "" , means not creating title row
        List<Row> rows = new ArrayList<>();
        rows.add(sheet.createRow(finishedLine));

        int[] headerRow = createHeaderRow(rowData, new int[]{finishedLine, 0});
        finishedLine += headerRow[0];
        return this;
    }

    private int[] createHeaderRow(List<HeaderElement> rowData, int[] rowAndColumn) {

        // re-order before create sheet rows
        for (HeaderElement headerElement : rowData) {
            if (headerElement.subElements == null || headerElement.subElements.size() == 0 || headerElement.order != Integer.MAX_VALUE)
                continue;
            Collections.sort(headerElement.subElements);
            headerElement.order = headerElement.subElements.get(0).order;
        }
        Collections.sort(rowData);

        int maxHtight = rowAndColumn[0];
        for (HeaderElement headerElement : rowData) {
            // create cells
            headerElement.recordRow = rowAndColumn[0];
            headerElement.recordColumn = rowAndColumn[1];
            if (headerElement.width != 0)
                sheet.setColumnWidth(rowAndColumn[1], headerElement.width * 256); // also set column width
            String fieldName = headerElement.generateCahinFieldName();
            if (StringUtils.isNotEmpty(fieldName))
                properties.put(PoiBaseConstants.VAR_COL + fieldName, Character.toString(((char) ('A' + rowAndColumn[1])))); // column variable todo
            createStringCell(getOrCreateRow(rowAndColumn[0]), rowAndColumn[1], headerElement.value);
            maxHtight = rowAndColumn[0] + 1;
            int[] rowAndColumn2 = new int[]{rowAndColumn[0], rowAndColumn[1] + 1}; // column add 1
            if (headerElement.subElements != null) {
                headerElement.setParentElementForChild();
                rowAndColumn2 = createHeaderRow(headerElement.subElements, new int[]{rowAndColumn[0] + 1, rowAndColumn[1]});
            }
            rowAndColumn[1] = rowAndColumn2[1];
            // horizonal merge
            try { // vertical merge
                sheet.addMergedRegion(new CellRangeAddress(headerElement.recordRow, headerElement.recordRow,
                        headerElement.recordColumn, rowAndColumn[1] - 1));
            } catch (IllegalArgumentException e) {
                LOGGER.error("合并单元格错误日志：" + e.getMessage());
                e.fillInStackTrace();
            }

            maxHtight = rowAndColumn2[0] > maxHtight ? rowAndColumn2[0] : maxHtight;
        }

        //vertical merge
        for (HeaderElement headerElement : rowData) {
            if (headerElement.subElements != null && headerElement.subElements.size() != 0) continue;
            try { // vertical merge
                sheet.addMergedRegion(new CellRangeAddress(headerElement.recordRow, maxHtight - 1,
                        headerElement.recordColumn, headerElement.recordColumn));
            } catch (IllegalArgumentException e) {
                LOGGER.error("合并单元格错误日志：" + e.getMessage());
                e.fillInStackTrace();
            }
        }
        return new int[]{maxHtight, rowAndColumn[1]};
    }

    private Row getOrCreateRow(int n) {
//        if (n-1 - rows.size() > 1) throw new RuntimeException("Wrong use of the method");
        if (n + 1 - rows.size() > 0) {
            for (int i = rows.size(); i <= n; i++) {
                rows.add(sheet.createRow(rows.size()));
            }
        }
        return rows.get(n);
    }

    public void insertOneMergedRow(List<ObjElement> rowData) {
        properties.put(VAR_PARENT_ROW_LIST, String.valueOf(finishedLine+1));
        properties.put(VAR_ROW_SUBLIST, String.valueOf(finishedLine+1));
//        VAR_ROW_SUBLIST
        rowDatas.add(rowData);
        // read one row or multiple rows
        int startColumn = 0;
//        finishedLine++;
        int maxRow = finishedLine;

        for (ObjElement valueElement : rowData) {
            if (valueElement.subObjList != null) {
                int[] rowAndColumn2 = new int[]{finishedLine, startColumn};  // or create a new one

                for (List<ObjElement> objElementList2 : valueElement.subObjList) {
                    rowAndColumn2 = insertSubObjList(rows, new int[]{rowAndColumn2[0], startColumn}, objElementList2);
                }
                startColumn = rowAndColumn2[1];
                valueElement.mergedRows = rowAndColumn2[0] - 1;
                maxRow = maxRow > rowAndColumn2[0] ? maxRow : rowAndColumn2[0];
                continue;
            }

            properties.put(VAR_ROW_LIST, String.valueOf(finishedLine+1));
            if (valueElement.type == 1) {
                createStringCell(getOrCreateRow(finishedLine), startColumn, valueElement.value);
            } else if (valueElement.type == 4) { // numerical
                createNumericCell(getOrCreateRow(finishedLine), startColumn, valueElement.value);
            } else if (valueElement.type == 3) { // formula
                createFormulaCell(getOrCreateRow(finishedLine), startColumn, tokenParser.parse(valueElement.value));
            } else { // create picture case
                try {

                    createImageCell(getDrawaing(), getOrCreateRow(finishedLine), startColumn, valueElement.value);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            valueElement.recordRow = finishedLine;
            valueElement.recordColumn = startColumn;
            startColumn++;

        }
        finishedLine = maxRow > finishedLine ? maxRow : finishedLine + 1;

        // vertically merge some columns by max mergedRows in the same level
        for (ObjElement valueElement : rowData) {
            if (valueElement.subObjList != null) {
                continue;
            }
            try { // vertical merge
                sheet.addMergedRegion(new CellRangeAddress(valueElement.recordRow, finishedLine - 1,
                        valueElement.recordColumn, valueElement.recordColumn));
            } catch (IllegalArgumentException e) {
                LOGGER.error("合并单元格错误日志：" + e.getMessage());
                e.fillInStackTrace();
            }
        }

    }


    private int[] insertSubObjList(List<Row> rows, int[] rowAndColumn, List<ObjElement> valueElement) { // int startRow, int startColumn int[2]{,}
        int[] rowAndColumn1 = new int[]{rowAndColumn[0], rowAndColumn[1]};  // or create a new one
        int maxRow = rowAndColumn[0];
        for (ObjElement objElement : valueElement) {
            if (objElement.subObjList != null) {
                int[] rowAndColumn2 = rowAndColumn;  // or create a new one
                String backUp = (String) properties.get(VAR_ROW_SUBLIST);
                properties.put(VAR_ROW_SUBLIST, String.valueOf(rowAndColumn2[0]+1));
                for (List<ObjElement> objElementList2 : objElement.subObjList) {
                    rowAndColumn2 = insertSubObjList(rows, new int[]{rowAndColumn2[0], rowAndColumn[1]}, objElementList2);
                }
                properties.put(VAR_ROW_SUBLIST, backUp);
                rowAndColumn1[1] = rowAndColumn2[1];
                objElement.mergedRows = rowAndColumn2[0] - 1;
                maxRow = maxRow > rowAndColumn2[0] ? maxRow : rowAndColumn2[0];
                continue;
            }
            properties.put(VAR_ROW_LIST, String.valueOf(finishedLine+1));
            if (objElement.type == 1) {
                createStringCell(getOrCreateRow(rowAndColumn[0]), rowAndColumn[1], objElement.value);
            } else if (objElement.type == 4) { // numerical
                createNumericCell(getOrCreateRow(rowAndColumn[0]), rowAndColumn[1], objElement.value);
            } else if (objElement.type == 3) { // formula
                createFormulaCell(getOrCreateRow(rowAndColumn[0]), rowAndColumn[1], tokenParser.parse(objElement.value));
            } else { // create picture case
                try {
                    createImageCell(getDrawaing(), getOrCreateRow(rowAndColumn[0]), rowAndColumn[1], objElement.value);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            objElement.recordRow = rowAndColumn[0];
            objElement.recordColumn = rowAndColumn[1];
            rowAndColumn[1]++;
        }
        rowAndColumn[0] = maxRow > rowAndColumn[0] + 1 ? maxRow : rowAndColumn[0] + 1;
        for (ObjElement valueElement2 : valueElement) {
            if (valueElement2.subObjList != null) {
                continue;
            }
            try { // vertical merge
                sheet.addMergedRegion(new CellRangeAddress(valueElement2.recordRow, rowAndColumn[0] - 1,
                        valueElement2.recordColumn, valueElement2.recordColumn));
            } catch (IllegalArgumentException e) {
                LOGGER.error("合并单元格错误日志：" + e.getMessage());
                e.fillInStackTrace();
            }
        }

        return rowAndColumn;

    }

    // port parameters
    public static class HeaderElement implements Comparable<HeaderElement> {

        private String value;
        private String groupName;
        private HeaderElement parentElementForField;
        private List<HeaderElement> subElements = null;
        private int recordRow;
        private int recordColumn;
        private String fieldName;
        private int order;
        private int width;

        public HeaderElement(String value, String groupName, int order, int width) {
            this.value = value;
            this.groupName = groupName;
            this.order = order;
            this.width = width;
        }

        public HeaderElement(String value, String groupName, int order, int width, String fieldName) {
            this.value = value;
            this.groupName = groupName;
            this.order = order;
            this.width = width;
            this.fieldName = fieldName;
        }

        public HeaderElement(String value, int order) {
            this.value = value;
            this.order = order;
        }

        public void addElements(List<HeaderElement> subElements) {
            if (subElements == null || subElements.size() == 0) return;
            if (this.subElements == null) this.subElements = new ArrayList<>();
            this.subElements.addAll(subElements);
        }

        public void addElement(HeaderElement element) {
            if (element == null) return;
            if (subElements == null) subElements = new ArrayList<>();
            subElements.add(element);
        }

        public List<HeaderElement> getSubElements() {
            if (subElements == null) {
                this.subElements = new ArrayList<>();
            }
            return subElements;
        }

        public void setParentElement(HeaderElement parentElement) {
            this.parentElementForField = parentElement;
        }

        public void setParentElementForChild() {
            if (this.subElements != null && this.subElements.size() > 0) {
                for (HeaderElement subElement : this.subElements) {
                    subElement.setParentElement(this);
                }
            }
        }

        public String generateCahinFieldName() {
            if (StringUtils.isEmpty(this.fieldName)) return "";
            StringBuilder stringBuilder = new StringBuilder(this.fieldName);
            HeaderElement parentElement = this.parentElementForField;
            while (parentElement != null && StringUtils.isEmpty(parentElement.fieldName)) {
                stringBuilder.insert(0, ".").insert(0, parentElement.fieldName);
            }
            return stringBuilder.toString();
        }

        @Override
        public int compareTo(HeaderElement prev) {
            return this.order - prev.order;
        }

    }

    public static class ObjElement {
        private int mergedRows = -1; // the rows that  subElements occupy , which will be used to merge vertically the parent cell
        private String value; // this value is not used if subElements is non-null
        private int type = 1; // 1 String , 4 Numeric, 3 Formula, otherwise image
        private List<List<ObjElement>> subObjList = null; // List<List<ObjElement>>
        private int recordRow;
        private int recordColumn;

        public ObjElement(String value, int type) {
            this.value = value;
            this.type = type;
        }

        public ObjElement(String value) {
            this.value = value;
        }

        public void addSubObj(List<ObjElement> subObj) {
            if (subObj == null || subObj.size() == 0) return;
            if (this.subObjList == null) this.subObjList = new ArrayList<>();
            this.subObjList.add(subObj);
        }

    }

    //update-begin--Author:xuelin  Date:20171018 for：TASK #2372 【excel】AutoPoi 导出类型，type增加数字类型--------------------
    private void createNumericCell(Row row, int index, String text, CellStyle style, ExcelExportServer.ExcelExportEntity entity) {
        Cell cell = row.createCell(index);
        if (StringUtils.isEmpty(text)) {
            cell.setCellValue("");
            cell.setCellType(CellType.BLANK);
        } else {
//			cell.setCellFormula(text);
            cell.setCellValue(Double.parseDouble(text));
            cell.setCellType(CellType.NUMERIC);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private void createFormulaCell(Row row, int index, String text, CellStyle style, ExcelExportServer.ExcelExportEntity entity) {
        Cell cell = row.createCell(index);
        if (StringUtils.isEmpty(text)) {
            cell.setCellValue("");
            cell.setCellType(CellType.BLANK);
        } else {
            cell.setCellFormula(text);
//			cell.setCellValue(Double.parseDouble(text));
//			cell.setCellType(CellType.FORMULA);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private void createFormulaCell(Row row, int index, String text) {
        Cell cell = row.createCell(index);
        cell.setCellStyle(cellStyle);
        if (StringUtils.isEmpty(text)) {
            cell.setCellValue("");
//            cell.setCellType(CellType.BLANK); // deprecated
        } else {
            cell.setCellFormula(text);
//			cell.setCellValue(Double.parseDouble(text));
//			cell.setCellType(CellType.FORMULA);
        }
    }

    private void createStringCell(Row row, int index, String text) {
        Cell cell = row.createCell(index);
        cell.setCellStyle(cellStyle);
        RichTextString Rtext;
        if (type.equals(ExcelType.HSSF)) {
            Rtext = new HSSFRichTextString(text);
        } else {
            Rtext = new XSSFRichTextString(text);
        }
        cell.setCellValue(Rtext);
    }


    private void createNumericCell(Row row, int index, String text) {
        Cell cell = row.createCell(index);
        if (StringUtils.isEmpty(text)) {
            cell.setCellValue("");
            cell.setCellType(CellType.BLANK);
        } else {
//			cell.setCellFormula(text);
            cell.setCellValue(Double.parseDouble(text));
            cell.setCellType(CellType.NUMERIC);
        }
    }

    private void createStringCell(Row row, int index, String text, CellStyle style) {
        Cell cell = row.createCell(index);
        cell.setCellStyle(cellStyle);
        if (style != null && style.getDataFormat() > 0 && style.getDataFormat() < 12) {
            cell.setCellValue(Double.parseDouble(text));
            cell.setCellType(CellType.NUMERIC);
        } else {
            RichTextString Rtext;
            if (type.equals(ExcelType.HSSF)) {
                Rtext = new HSSFRichTextString(text);
            } else {
                Rtext = new XSSFRichTextString(text);
            }
            cell.setCellValue(Rtext);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }

    }


}
