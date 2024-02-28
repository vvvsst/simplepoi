
package org.simplepoi.excel.export;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.simplepoi.excel.ExportParams;
import org.simplepoi.excel.ReflectionUtil;
import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelEntity;
import org.simplepoi.excel.annotation.ExcelField;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.excel.exception.ExcelExportException;
import org.simplepoi.excel.constant.ExcelExportEnum;
import org.simplepoi.excel.imports.ExcelImportServer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class ExcelExportServer {
    protected ExcelType type = ExcelType.HSSF;

    private Map<Integer, Double> statistics = new HashMap<Integer, Double>();

    private static final DecimalFormat DOUBLE_FORMAT = new DecimalFormat("######0.00");
    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelExportServer.class);

    // 最大行数,超过自动多Sheet
    private int MAX_NUM = 60000;

    public void createSheet(Workbook workbook, ExportParams entity, Class<?> pojoClass, Collection<?> dataSet, String[] exportFields) {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel export start ,class is {}", pojoClass);
            LOGGER.debug("Excel version is {}", entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || pojoClass == null || dataSet == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        type = entity.getType();
        if (type.equals(ExcelType.XSSF)) {
            MAX_NUM = 1000000;
        }
        Sheet sheet = null;
        try {
            sheet = workbook.createSheet(entity.getSheetName());
        } catch (Exception e) {
            // 重复遍历,出现了重名现象,创建非指定的名称Sheet
            sheet = workbook.createSheet();
        }
        List<ExcelSheetManipulator.HeaderElement> headerElements = new ArrayList<>(); // data of a format that will be used to create sheets
        List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();

        ExcelSheetManipulator manipulator = new ExcelSheetManipulator(sheet, type); // actually used to create sheet cells

        try {
            // 得到所有字段
            Field[] fields = ReflectionUtil.getClassFields(pojoClass);
            //支持自定义导出字段
            if (exportFields != null) {
                List<Field> list = new ArrayList<Field>(Arrays.asList(fields));
                for (int i = 0; i < list.size(); i++) { // one-to-many case @ExcelCollection not considered
                    if (!Arrays.asList(exportFields).contains(list.get(i).getName())) {
                        list.remove(i);
                        i--;
                    }
                }

                if (list.size() > 0) {
                    fields = list.toArray(new Field[0]);
                } else {
                    fields = null;
                }
            }

            assert fields != null;
            readAllExcelFields(fields, excelParams, pojoClass, headerElements); // 获得@Excel注解的值，excelParams,
            manipulator.createTitleAndHeaderRow(headerElements, null, null); // null means do not create title row

            Iterator<?> its = dataSet.iterator();
            List<Object> tempList = new ArrayList<Object>();
            while (its.hasNext()) { // creat sheet for every row or every element in the list
                Object t = its.next();
                List<ExcelSheetManipulator.ObjElement> objElements = new ArrayList<>();
                convertToSheetDataFormat(t, excelParams, objElements);
                manipulator.insertOneMergedRow(objElements);
                tempList.add(t);
                if (manipulator.getFinishedLine() >= MAX_NUM) break;
            }

            if (entity.getFreezeCol() != 0) {
                sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
            }

            its = dataSet.iterator();
            for (int i = 0, le = tempList.size(); i < le; i++) {
                its.next();
                its.remove();
            }

            // 发现还有剩余list 继续循环创建Sheet, rows number reach the MAX_NUM, needs another sheet
            if (dataSet.size() > 0) { // recursive call
                createSheet(workbook, entity, pojoClass, dataSet, exportFields);
            }

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    private Properties replenishProperties(int index, List<ExcelExportEntity> excelParams) { // index represents row number
        Properties prop = new Properties();
        char columnLetter = 'A';
        for (ExcelExportEntity excelParam : excelParams) {
            if (excelParam.getList() != null) {
            } else {
                // to be modified todo
//                prop.setProperty(excelParams.get(i).getKey().toString(), columnLetter + String.valueOf(index));
                columnLetter++;
            }
        }
        return prop;
    }

    /**
     * 合计统计信息
     *
     * @param index
     * @param text
     * @param entity
     */
    private void addStatisticsData(Integer index, String text, ExcelExportEntity entity) {
        if (entity != null && entity.isStatistics()) {
            Double temp = 0D;
            if (!statistics.containsKey(index)) {
                statistics.put(index, temp);
            }
            try {
                temp = Double.valueOf(text);
            } catch (NumberFormatException e) {
            }
            statistics.put(index, statistics.get(index) + temp);
        }
    }


    public void readAllExcelFields(Field[] fields, List<ExcelExportEntity> excelParams, Class<?> pojoClass,
                                   List<ExcelSheetManipulator.HeaderElement> headerElements) throws Exception {
        HashMap<String, List<ExcelSheetManipulator.HeaderElement>> groupHeaderListMap = new HashMap<>();
        HashMap<String, List<ExcelExportEntity>> groupExportListMap = new HashMap<>();
        for (Field field : fields) {
            ExcelField excelField1 = field.getAnnotation(ExcelField.class);
            ExcelCollection excelCollection = field.getAnnotation(ExcelCollection.class);
            ExcelEntity excelEntity = field.getAnnotation(ExcelEntity.class);
            ExcelSheetManipulator.HeaderElement headerElement = null;
            ExcelExportEntity excelExportEntity = null;
            if (excelCollection != null && ReflectionUtil.isCollection(field.getType())) {
                excelExportEntity = ExcelExportEntity.createFromExcelCollection(pojoClass, field, excelCollection);
                ParameterizedType pt = (ParameterizedType) field.getGenericType();
                Class<?> clz = (Class<?>) pt.getActualTypeArguments()[0];
                // SUB - headerElements should be given here
                headerElement = new ExcelSheetManipulator.HeaderElement(excelExportEntity.getName(), null,
                        excelExportEntity.getOrderNum(), excelExportEntity.getWidth(), field.getName());
                readAllExcelFields(ReflectionUtil.getClassFields(clz),
                        excelExportEntity.getList(), clz, headerElement.getSubElements());

            } else if (excelField1 != null && ReflectionUtil.isJavaClass(field) && excelField1.type() != 5) {
                if (StringUtils.isNotEmpty(excelField1.groupName())) {
                    List<ExcelExportEntity> excelExportList = groupExportListMap.putIfAbsent(excelField1.groupName(), new ArrayList<>());
                    if (excelExportList == null) excelExportList = groupExportListMap.get(excelField1.groupName());
                    ExcelExportEntity excelExportEntity1 = ExcelExportEntity.createExcelExportEntity(field, pojoClass);
                    excelExportList.add(excelExportEntity1);

                    List<ExcelSheetManipulator.HeaderElement> headerElements1 = groupHeaderListMap.putIfAbsent(excelExportEntity1.getGroupName(), new ArrayList<>());
                    if (headerElements1 == null) headerElements1 = groupHeaderListMap.get(excelField1.groupName());
                    headerElements1.add(new ExcelSheetManipulator.HeaderElement(excelExportEntity1.getName(), excelExportEntity1.getGroupName(),
                            excelExportEntity1.getOrderNum(), excelExportEntity1.getWidth(),field.getName()));
                } else {
                    excelExportEntity = ExcelExportEntity.createExcelExportEntity(field, pojoClass);
                    headerElement = new ExcelSheetManipulator.HeaderElement(excelExportEntity.getName(), excelExportEntity.getGroupName(),
                            excelExportEntity.getOrderNum(), excelExportEntity.getWidth(), field.getName());
                }

                // if not basic class, consider it as a user-defined class, fields of which should be basic javaClass to recursively obtain field information
            } else if (excelEntity != null) {

                headerElement = new ExcelSheetManipulator.HeaderElement(excelEntity.name(), null,
                        excelEntity.order(), 0, field.getName());
                Map<String, ExcelImportServer.ExcelImportEntity> subExcelParams = new HashMap<>();
                readAllExcelFields(ReflectionUtil.getClassFields(field.getType()), excelParams, field.getType(),
                        headerElement.getSubElements());
            }
            if (headerElement != null) headerElements.add(headerElement);
            if (excelExportEntity != null) excelParams.add(excelExportEntity);
        }
        for (String groupName : groupHeaderListMap.keySet()) {
            List<ExcelSheetManipulator.HeaderElement> headerElements1 = groupHeaderListMap.get(groupName);
            ExcelSheetManipulator.HeaderElement headerElement = new ExcelSheetManipulator.HeaderElement(groupName, Integer.MAX_VALUE); // order to be determined
            headerElement.addElements(headerElements1);
            headerElements.add(headerElement);
        }
        for (String groupName : groupExportListMap.keySet()) {
            List<ExcelExportEntity> excelExportEntities = groupExportListMap.get(groupName);
            ExcelExportEntity excelExport = new ExcelExportEntity();
            Collections.sort(excelExportEntities);
            excelExport.setOrderNum(excelExportEntities.get(0).getOrderNum());
            excelExport.setName(groupName);
            excelExport.getList().addAll(excelExportEntities);
            excelParams.add(excelExport);
        }

        // reorder according to order parameter, after a level of list has been added
        Collections.sort(excelParams);
    }


    //region frequently used

    private void convertToSheetDataFormat(Object t, List<ExcelExportEntity> excelParams, List<ExcelSheetManipulator.ObjElement> objElements) throws Exception {
        int index = 5;
//        Properties properties = replenishProperties(index + 1, excelParams); // formula function will be implemented in ExcelSheetManipulator todo
//        GenericTokenParser genericTokenParser = new GenericTokenParser("${", "}", properties);
        ExcelExportEntity entity;
        for (ExcelExportEntity excelParam : excelParams) {
            entity = excelParam;

            if (entity.getList() == null || entity.getList().size() == 0) { // ordinary field of pojo
                Object value = getCellValue(entity, t);
                ExcelSheetManipulator.ObjElement element = new ExcelSheetManipulator.ObjElement(value == null ? "" : value.toString(), entity.getType());
                objElements.add(element);
            } else if (entity.getList() != null && entity.getList().size() > 0 && entity.getMethod() == null) { // horizontal merge, group case, this case can be treated like the usual one, and will be removed todo
                convertToSheetDataFormat(t, entity.getList(), objElements);
            } else if (entity.getList() != null && entity.getList().size() > 0 && entity.getMethod() != null) {  // @ExcelCollection case
                Collection<?> list = getListCellValue(entity, t);
                ExcelSheetManipulator.ObjElement element = new ExcelSheetManipulator.ObjElement(null);
                for (Object obj : list) {
                    List<ExcelSheetManipulator.ObjElement> subobj = new ArrayList<>();
                    convertToSheetDataFormat(obj, entity.getList(), subobj);  // recursive call is needed here todo
                    element.addSubObj(subobj);
                }
                objElements.add(element);
            } else throw new RuntimeException(" unsupported ExcelExportEntity ");

        }
    }


    /**
     * 获取填如这个cell的值,提供一些附加功能
     *
     * @param entity
     * @param obj
     * @return
     * @throws Exception
     */
    private Object getCellValue(ExcelExportEntity entity, Object obj) throws Exception {
        if (entity.type == 3) return entity.getFormulaExpr(); // 公式只需要传入公式表达式
        Object value = entity.getMethod().invoke(obj, new Object[]{});
        value = Optional.ofNullable(value).orElse("");
        if (StringUtils.isEmpty(value.toString())) {
            return "";
        }
        if (StringUtils.isNotEmpty(entity.getNumFormat()) && value != null) {
            value = new DecimalFormat(entity.getNumFormat()).format(value);
        }
        if (StringUtils.isNotEmpty(entity.getFormat())) {
            value = formatValue(value, entity);
        }
        if (entity.getReplace() != null && entity.getReplace().length > 0) {
            if (value == null) {
                value = "";
            }
            String oldVal = value.toString();
            if (entity.isMultiReplace()) {
                value = multiReplaceValue(entity.getReplace(), String.valueOf(value));
            } else {
                value = replaceValue(entity.getReplace(), String.valueOf(value));
            }
            if (oldVal.equals(value)) {
            }
        }

        if (StringUtils.isNotEmpty(entity.getSuffix()) && value != null) {
            value = value + entity.getSuffix();
        }
        return value == null ? "" : value.toString();
    }

    /**
     * 获取集合的值
     *
     * @param entity
     * @param obj
     * @return
     * @throws Exception
     */
    private Collection<?> getListCellValue(ExcelExportEntity entity, Object obj) throws Exception {
        Object value;
        value = entity.getMethod().invoke(obj);
        if (value instanceof Collection) {
            return (Collection<?>) value;
        } else {
            List list = new ArrayList();
            list.add(value);
            return list;
        }
    }

    //update-begin-author:taoyan date：20180731 for:TASK #3038 【bug】Excel 导出多个值（逗号隔开的情况下，导出字典值是ID值）

    /**
     * 如果需要被替换的值是多选项，则每一项之间有逗号隔开，走以下方法
     *
     * @author taoYan
     * @since 2018年7月31日
     */
    private Object multiReplaceValue(String[] replace, String value) {
        if (value.indexOf(",") > 0) {
            String[] radioVals = value.split(",");
            String[] temp;
            String result = "";
            for (int i = 0; i < radioVals.length; i++) {
                String radio = radioVals[i];
                for (String str : replace) {
                    temp = str.split("_");
                    //update-begin-author:liusq date：20210127 for:字符串截取修改
                    temp = getValueArr(str);
                    //update-end-author:liusq date：20210127 for:字符串截取修改

                    //update-begin---author:scott   Date:20211220  for：[issues/I4MBB3]@Excel dicText字段的值有下划线时，导入功能不能正确解析---
                    if (radio.equals(temp[1]) || radio.replace("_", "---").equals(temp[1])) {
                        result = result.concat(temp[0]) + ",";
                        break;
                    }
                    //update-end---author:scott   Date:20211220  for：[issues/I4MBB3]@Excel dicText字段的值有下划线时，导入功能不能正确解析---
                }
            }
            if (result.equals("")) {
                result = value;
            } else {
                result = result.substring(0, result.length() - 1);
            }
            return result;
        } else {
            return replaceValue(replace, value);
        }
    }
    //update-end-author:taoyan date：20180731 for:TASK #3038 【bug】Excel 导出多个值（逗号隔开的情况下，导出字典值是ID值）


    private Object formatValue(Object value, ExcelExportEntity entity) throws Exception {
        Date temp = null;
        //update-begin-author:wangshuai date:20201118 for:Excel导出错误原因，value为""字符串，gitee I249JF
        if ("".equals(value)) {
            value = null;
        }
        //update-begin-author:wangshuai date:20201118 for:Excel导出错误原因，value为""字符串，gitee I249JF
        if (value instanceof String && entity.getDatabaseFormat() != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getDatabaseFormat());
            temp = format.parse(value.toString());
        } else if (value instanceof Date) {
            temp = (Date) value;
            //update-begin-author:taoyan date:2022-5-17 for: mybatis-plus升级 时间字段变成了jdk8的LocalDateTime，导致格式化失败
        } else if (value instanceof LocalDateTime) {
            LocalDateTime ldt = (LocalDateTime) value;
            DateTimeFormatter format = DateTimeFormatter.ofPattern(entity.getFormat());
            return format.format(ldt);
        } else if (value instanceof LocalDate) {
            LocalDate ld = (LocalDate) value;
            DateTimeFormatter format = DateTimeFormatter.ofPattern(entity.getFormat());
            return format.format(ld);
        }
        //update-end-author:taoyan date:2022-5-17 for: mybatis-plus升级 时间字段变成了jdk8的LocalDateTime，导致格式化失败
        if (temp != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            value = format.format(temp);
        }
        return value;
    }

    private Object replaceValue(String[] replace, String value) {
        String[] temp;
        for (String str : replace) {
            //temp = str.split("_"); {'男_sheng_1','女_2'}
            //update-begin-author:liusq date：20210127 for:字符串截取修改
            temp = getValueArr(str);
            //update-end-author:liusq date：20210127 for:字符串截取修改

            //update-begin---author:scott   Date:20211220  for：[issues/I4MBB3]@Excel dicText字段的值有下划线时，导入功能不能正确解析---
            if (value.equals(temp[1]) || value.replace("_", "---").equals(temp[1])) {
                value = temp[0];
                break;
            }
            //update-end---author:scott   Date:20211220  for：[issues/I4MBB3]@Excel dicText字段的值有下划线时，导入功能不能正确解析---
        }
        return value;
    }


    /**
     * 字典文本中含多个下划线横岗，取最后一个（解决空值情况）
     *
     * @param val
     * @return
     */
    private static String[] getValueArr(String val) {
        int i = val.lastIndexOf("_");//最后一个分隔符的位置
        String[] c = new String[2];
        c[0] = val.substring(0, i); //label
        c[1] = val.substring(i + 1); //key
        return c;
    }

    //endregion


    public static class ExcelExportEntity implements Comparable<ExcelExportEntity> {
        /**
         * 对应name
         */
        protected String name;

        /**
         * 对应 class 的 属性name
         */
        protected String fieldName;
        /**
         * 对应type
         */
        private int type = 1;
        /**
         * 公式表达式
         */
        private String formulaExpr;
        /**
         * 数据库格式
         */
        private String databaseFormat;
        /**
         * 导出日期格式
         */
        private String format;
        private String[] importFormats; // 导入日期所接受的格式

        /**
         * 数字格式化,参数是Pattern,使用的对象是DecimalFormat
         */
        private String numFormat;
        /**
         * 替换值表达式 ："男_1","女_0"
         */
        private String[] replace;
        /**
         * 替换是否是替换多个值
         */
        private boolean multiReplace;
        /**
         * set/get方法, 导出 为 get
         */
        private Method method;
        /**
         * 固定的列
         */
        private Integer fixedIndex;


        public String getDatabaseFormat() {
            return databaseFormat;
        }

        public String getFormat() {
            return format;
        }

        public String[] getImportFormats() {
            return importFormats;
        }

        public Method getMethod() {
            return method;
        }


        public String getName() {
            return name;
        }

        public String[] getReplace() {
            return replace;
        }

        public int getType() {
            return type;
        }

        public String getFormulaExpr() {
            return formulaExpr;
        }


        public void setDatabaseFormat(String databaseFormat) {
            this.databaseFormat = databaseFormat;
        }

        public void setFormat(String format) {
            this.format = format;
        }

        public void setImportFormats(String[] format) {
            this.importFormats = format;
        }

        public void setMethod(Method method) {
            this.method = method;
        }


        public void setName(String name) {
            this.name = name;
        }

        public void setReplace(String[] replace) {
            this.replace = replace;
        }

        public void setType(int type) {
            this.type = type;
        }

        public void setFormulaExpr(String formulaExpr) {
            this.formulaExpr = formulaExpr;
        }

        public boolean isMultiReplace() {
            return multiReplace;
        }

        public void setMultiReplace(boolean multiReplace) {
            this.multiReplace = multiReplace;
        }

        public String getNumFormat() {
            return numFormat;
        }

        public void setNumFormat(String numFormat) {
            this.numFormat = numFormat;
        }

        public Integer getFixedIndex() {
            return fixedIndex;
        }

        public void setFixedIndex(Integer fixedIndex) {
            this.fixedIndex = fixedIndex;
        }


        private int width = 10;

        private double height = 10;

        /**
         * 图片的类型,1是文件地址(class目录),2是数据库字节,3是文件地址(磁盘目录)，4网络图片
         */
        private int exportImageType = 3;

        /**
         * 排序顺序
         */
        private int orderNum = 0;

        /**
         * 后缀
         */
        private String suffix;
        /**
         * 统计
         */
        private boolean isStatistics;

        /**
         * 父表头的名称
         */
        private String groupName; // used to construct headerElemtn

        // an inital value is given, to derteim is it annotated with @ExcelCollection, use its size==0
        private List<ExcelExportEntity> list = new ArrayList<>();

        public ExcelExportEntity() {
        }

        public ExcelExportEntity(String name) {
            this.name = name;
        }

        public int getExportImageType() {
            return exportImageType;
        }

        public double getHeight() {
            return height;
        }


        public List<ExcelExportEntity> getList() {
            return list;
        }

        public int getOrderNum() {
            return orderNum;
        }

        public int getWidth() {
            return width;
        }

        public void setExportImageType(int exportImageType) {
            this.exportImageType = exportImageType;
        }

        public void setHeight(double height) {
            this.height = height;
        }


        public void setList(List<ExcelExportEntity> list) {
            this.list = list;
        }

        public void setOrderNum(int orderNum) {
            this.orderNum = orderNum;
        }

        public void setWidth(int width) {
            this.width = width;
        }

        public String getSuffix() {
            return suffix;
        }

        public void setSuffix(String suffix) {
            this.suffix = suffix;
        }

        public boolean isStatistics() {
            return isStatistics;
        }

        public void setStatistics(boolean isStatistics) {
            this.isStatistics = isStatistics;
        }


        public String getGroupName() {
            return groupName;
        }

        public void setGroupName(String groupName) {
            this.groupName = groupName;
        }


        @Override
        public int compareTo(ExcelExportEntity prev) {
            return this.getOrderNum() - prev.getOrderNum();
        }


        private static ExcelExportEntity createFromExcelCollection(Class<?> pojoClass, Field field, ExcelCollection excelCollection) {

            List<ExcelExportEntity> list = new ArrayList<ExcelExportEntity>();
            // SUB - headerElements should be given here todo
            ExcelExportEntity excelExportEntity = new ExcelExportEntity();
            excelExportEntity.setName(excelCollection.name());
            excelExportEntity.setOrderNum(Integer.parseInt(excelCollection.orderNum()));
            try {
                excelExportEntity.setMethod(ReflectionUtil.getGetMethod(field.getName(), pojoClass));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            excelExportEntity.setList(list);
            return excelExportEntity;
        }

        /**
         * 创建导出实体对象
         *
         * @param field
         * @param pojoClass
         * @return
         * @throws Exception
         */
        private static ExcelExportEntity createExcelExportEntity(Field field, Class<?> pojoClass) throws Exception {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            ExcelExportEntity excelEntity = new ExcelExportEntity();
            excelEntity.setType(excelField.type());
            excelEntity.setFormulaExpr(excelField.formulaExpr());
            excelEntity.fieldName = field.getName();
            excelEntity.setName(excelField.name());
            excelEntity.setWidth(excelField.width());
            excelEntity.setHeight(excelField.height());
            excelEntity.setReplace(excelField.replace());

            excelEntity.setOrderNum(Integer.parseInt(excelField.orderNum()));
            excelEntity.setExportImageType(excelField.imageType());
            excelEntity.setSuffix(excelField.suffix());
            excelEntity.setDatabaseFormat(excelField.databaseFormat());
            excelEntity.setFormat(StringUtils.isNotEmpty(excelField.exportFormat()) ? excelField.exportFormat() : excelField.format());
            excelEntity.setStatistics(excelField.isStatistics());
            excelEntity.setNumFormat(excelField.numFormat());
            excelEntity.setMethod(ReflectionUtil.getGetMethod(field.getName(), pojoClass, excelField.exportConvert()));
            excelEntity.setMultiReplace(excelField.multiReplace());
            excelEntity.setGroupName(excelField.groupName());
            return excelEntity;
        }
    }
}
