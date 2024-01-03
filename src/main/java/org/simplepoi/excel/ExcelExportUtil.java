/**
 * Copyright 2013-2015 JEECG (jeecgos@163.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.simplepoi.excel;

import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.excel.export.ExcelExportServer;

/**
 * excel 导出工具类
 *
 * @author JEECG
 * @version 1.0
 * @date 2013-10-17
 */
public final class ExcelExportUtil {
    //单sheet最大值
    public static int USE_SXSSF_LIMIT = 100000;

    private ExcelExportUtil() {
    }


    /**
     * 根据Entity创建对应的Excel
     *
     * @param entity
     *            表格标题属性
     * @param pojoClass
     *            Excel对象Class
     * @param dataSet
     *            Excel对象数据List
     * @param exportFields
     * 	          自定义导出Excel字段数组
     */
    public static Workbook exportExcel(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet, String[] exportFields) {
        Workbook workbook;
        if (ExcelType.HSSF.equals(entity.getType())) {
            workbook = new HSSFWorkbook();
        } else if (dataSet.size() < 1000) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new SXSSFWorkbook();
        }
        new ExcelExportServer().createSheet(workbook, entity, pojoClass, dataSet, exportFields);
        return workbook;
    }
    //---update-end-----autor:scott------date:20191016-------for:导出字段支持自定义--------


    /**
     * 根据Entity创建对应的Excel
     *
     * @param entity
     *            表格标题属性
     * @param pojoClass
     *            Excel对象Class
     * @param dataSet
     *            Excel对象数据List
     */
    public static Workbook exportExcel(ExportParams entity, Class<?> pojoClass, Collection<?> dataSet) {
        Workbook workbook;
        if (ExcelType.HSSF.equals(entity.getType())) {
            workbook = new HSSFWorkbook();
        } else if (dataSet.size() < 1000) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new SXSSFWorkbook();
        }
        new ExcelExportServer().createSheet(workbook, entity, pojoClass, dataSet, null);
        return workbook;
    }



    /**
     * 一个excel 创建多个sheet
     *
     * @param list
     *            多个Map key title 对应表格Title key entity 对应表格对应实体 key data
     *            Collection 数据
     * @return
     */
    public static Workbook exportExcel(List<Map<String, Object>> list, ExcelType type) {
        Workbook workbook;
        if (ExcelType.HSSF.equals(type)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        for (Map<String, Object> map : list) {
            String[] exportFields = null;
            Object exportFieldStr = map.get("exportFields"); // exportFields  NormalExcelConstants.EXPORT_FIELDS
            if (exportFieldStr != null && exportFieldStr != "")
                exportFields = exportFieldStr.toString().split(",");
            ExcelExportServer server = new ExcelExportServer();
            server.createSheet(workbook, (ExportParams) map.get("title"), (Class<?>) map.get("entity"), (Collection<?>) map.get("data"), exportFields);
        }
        return workbook;
    }

}
