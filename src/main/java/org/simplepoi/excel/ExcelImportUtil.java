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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.simplepoi.excel.imports.ExcelImportServer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel 导入工具
 *
 * @author JEECG
 * @date 2013-9-24
 * @version 1.0
 */
@SuppressWarnings({"unchecked"})
public final class ExcelImportUtil {

    private ExcelImportUtil() {
    }

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelImportUtil.class);

    /**
     * Excel 导入 数据源本地文件,不返回校验结果 导入 字 段类型 Integer,Long,Double,Date,String,Boolean
     *
     * @param file
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcel(File file, Class<T> pojoClass, ImportParams params) {
        FileInputStream in = null;
        List<T> result = null;
        try {
            in = new FileInputStream(file);
            result = new ExcelImportServer().importExcel(in, pojoClass, params);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                LOGGER.error(e.getMessage(), e);
            }
        }
        return result;
    }

    /**
     * Excel 导入 数据源IO流,不返回校验结果 导入 字段类型 Integer,Long,Double,Date,String,Boolean
     *
     * @param pojoClass
     * @param params
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcel(InputStream inputstream, Class<T> pojoClass, ImportParams params) throws Exception {
        return new ExcelImportServer().importExcel(inputstream, pojoClass, params);
    }

    public static <T> List<T> importExcelFromDesktop(Class<T> pojoClass, String filename) throws Exception {
        ImportParams params = new ImportParams();
        return importExcelFromDesktop(pojoClass, params, filename);
    }

    public static <T> List<T> importExcelFromDesktop(Class<T> pojoClass, ImportParams params, String filename) throws Exception {
        String userHome = System.getProperty("user.home") + "\\Desktop\\" + filename;
        File file = new File(userHome);
        FileInputStream fileInputStream = new FileInputStream(file);
        return new ExcelImportServer().importExcel(fileInputStream, pojoClass, params);
    }


}
