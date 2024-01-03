package org.simplepoi.excel.imports;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.util.Date;

import static org.simplepoi.excel.ExelCommonUtil.trimAllWhitespace;


public class CellValueServer {

    private static final Logger LOGGER = LoggerFactory.getLogger(CellValueServer.class);

    private final DataFormatter numFormatter = new DataFormatter();
    final private ExcelPropertyEditor numEditor = new NumberEditor();

    //update-begin-author:taoyan date:20180807 for:导入多值替换--

    /**
     * 单值替换 ,若没找到则原值返回, 返回 null
     */
    private String replaceSingleValue(String[] replace, String temp) {
        String[] tempArr;
        for (int i = 0; i < replace.length; i++) {
            //update-begin---author:scott   Date:20211220  for：[issues/I4MBB3]@Excel dicText字段的值有下划线时，导入功能不能正确解析---
            //tempArr = replace[i].split("_");
            tempArr = getValueArr(replace[i]);
            if (temp.equals(tempArr[0]) || temp.replace("_", "---").equals(tempArr[0])) {
                //update-begin---author:wangshuai ---date:20220422  for：导入字典替换需要将---替换成_，不然数据库会存--- ------------
                if (tempArr[1].contains("---")) {
                    return tempArr[1].replace("---", "_");
                }
                //update-end---author:wangshuai ---date:20220422  for：导入字典替换需要将---替换成_，不然数据库会存--- --------------
                return tempArr[1];
            }
            //update-end---author:scott   Date:20211220  for：[issues/I4MBB3]@Excel dicText字段的值有下划线时，导入功能不能正确解析---
        }
        return "";
//		return temp;
    }
    //update-end-author:taoyan date:20180807 for:导入多值替换--


    /**
     * 字典文本中含多个下划线横岗，取最后一个（解决空值情况）
     *
     * @param val
     * @return
     */
    private String[] getValueArr(String val) {
        int i = val.lastIndexOf("_");//最后一个分隔符的位置
        String[] c = new String[2];
        c[0] = val.substring(0, i); //label
        c[1] = val.substring(i + 1); //key
        return c;
    }

    //region Frequently used


    /**
     * 导入支持多值替换
     *
     * @param replace      数据库中字典查询出来的数组
     * @param result       excel单元格获取的值
     * @param multiReplace 是否支持多值替换
     * @author taoYan
     * @since 2018年8月7日
     */
    private Object replaceValue(String[] replace, Object result, boolean multiReplace) {
        if (result == null) {
            return "";
        }
        if (replace == null || replace.length <= 0) {
            return result;
        }
        String temp = String.valueOf(result);
        String backValue = "";
        if (temp.indexOf(",") > 0 && multiReplace) {
            //原值中带有逗号，认为他是多值的
            String[] multiReplaces = temp.split(",");
            for (String str : multiReplaces) {
                backValue = backValue.concat(replaceSingleValue(replace, str) + ",");
            }
            if (backValue.equals("")) {
                backValue = temp;
            } else {
                backValue = backValue.substring(0, backValue.length() - 1);
            }
        } else {
            backValue = replaceSingleValue(replace, temp);
        }
        //update-begin-author:liusq date:20210204 for:字典替换失败提示日志
        if (backValue.equals(temp) || backValue.equals("")) {
            LOGGER.warn("===========替换失败,替换值:{},要转换的导入值:{}==========", replace, temp);
        }
        //update-end-author:liusq date:20210204 for:字典替换失败提示日志
        return backValue;
    }


    public Object getValue(Object object, Cell cell, ExcelImportServer.ExcelImportEntity entity) throws Exception {
        Class<?> fieldType = entity.getClassType2();
        // the final type is determined by class to be set of entity field (typeA) and type of Excel cell (typeB)
        // the type is divided into three classes, date, string, number. or unsupported type, throw exception/ return null
        // tyepA vs typeB
        // string vs number with format : get the formatted number
        // string vs string with format : get the formatted string
        if (fieldType == String.class) {
            if (cell.getCellType() == CellType.NUMERIC)
                return trimAllWhitespace(numFormatter.formatCellValue(cell));
            return getFormattedCellStr(cell);
        }
        // number vs number with format : don't consider format for this case, needs editor to set property from double to Decimal/Integer/...
        if (cell.getCellType() == CellType.NUMERIC && numEditor.supports(Double.class, fieldType)) {
            return numEditor.convert(cell.getNumericCellValue(), fieldType);
        }

        // convertToNumber
        // number/decimal vs string with format : not consider format for this case, try to convert it to number with value of
        // number(replace) vs string format : use the formated string to get the number or not formated one
        if (cell.getCellType() == CellType.STRING && numEditor.supports(String.class, fieldType)) {
            String[] replace = entity.getReplace();
            if (replace != null && replace.length > 0) {
                Object value = replaceValue(replace, getFormattedCellStr(cell), entity.isMultiReplace());
                return numEditor.convert(value, fieldType);
            }
            return numEditor.convert(cell.getStringCellValue(), fieldType);
        }


        // convertToDate
        // date  vs number/date with format : convert with the format
        // date  vs string with format :  your own converter, using multiple date format in @Excel , add the parameter
        ExcelPropertyEditor dateEditor;
        if (entity.getImportFormats() != null && entity.getImportFormats().length > 0)
            dateEditor = new DateEditor(entity.getImportFormats());
        else dateEditor = new DateEditor(entity.getFormat());
        if (cell.getCellType() == CellType.NUMERIC && dateEditor.supports(Date.class, fieldType)) {
            return dateEditor.convert(cell.getDateCellValue(), fieldType);
        }
        if (cell.getCellType() == CellType.STRING && dateEditor.supports(String.class, fieldType)) {
            // dateEditor
            return dateEditor.convert(getFormattedCellStr(cell), fieldType);
        }
        return null;
    }

    private String getFormattedCellStr(Cell cell) throws InvocationTargetException, IllegalAccessException {
        ExcelTextFormatter textFormatter = ExcelTextFormatter.getInstance();
        String format = textFormatter.format(cell.getStringCellValue(), cell.getCellStyle().getDataFormatString());
        return trimAllWhitespace(format);
    }

    //endregion

}
