
package org.simplepoi.excel.imports;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.simplepoi.excel.ImportParams;
import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelEntity;
import org.simplepoi.excel.annotation.ExcelField;
import org.simplepoi.excel.ReflectionUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import static org.simplepoi.excel.ReflectionUtil.getSetMethod;
import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIELD;
import static org.simplepoi.excel.ReflectionUtil.createObject;
import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;
import static org.simplepoi.excel.imports.ImportImageSupport.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.*;


public class ExcelImportServer {

    private final static Logger LOGGER = LoggerFactory.getLogger(ExcelImportServer.class);
    private final CellValueServer cellValueServer = new CellValueServer();
    private Map<Integer, String> titleMap = new HashMap<>();
    private Map<String, PictureData> pictures = new HashMap<>();
    private ImportParams params = null;
    private final Map<String, ExcelImportEntity> excelParams = new HashMap<>(); // correspond to every field to be exported
    private final List<ExcelCollectionParams> excelCollection = new ArrayList<>(); // ExcelImportEntity is constructed from annotation Excel


    public ExcelImportServer() {
    }


    //region Frequently used

    /**
     * Excel 导入 field 字段类型 Integer,Long,Double,Date,String,Boolean
     */
    public <T> List<T> importExcel(InputStream inputstream, Class<T> pojoClass, ImportParams params) throws Exception {
        this.params = params;
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel import start ,class is {}", pojoClass);
        }
        List<T> result = new ArrayList<>();
        Workbook book;
        boolean isXSSFWorkbook = false;
        if (!(inputstream.markSupported())) {
            inputstream = new PushbackInputStream(inputstream, 8);
        }
        book = WorkbookFactory.create(inputstream);
        if (book instanceof XSSFWorkbook) {
            isXSSFWorkbook = true;
        }
        if (params.getSheetNum() == 0) { // 多sheet导入改造点 获取导入文本的sheet数
            int sheetNum = book.getNumberOfSheets();
            if (sheetNum > 0) { //update-end-author:taoyan date:20211210 for:https://gitee.com/jeecg/jeecg-boot/issues/I45C32 导入空白sheet报错
                params.setSheetNum(sheetNum);
            }
        }
        for (int i = params.getStartSheetIndex(); i < params.getStartSheetIndex() // getStartSheetIndex 开始读取sheet的位置
                + params.getSheetNum(); i++) { // getSheetNum 需要读取的 sheet 数量

            if (isXSSFWorkbook) {
                pictures = getSheetPictrues07((XSSFSheet) book.getSheetAt(i), (XSSFWorkbook) book);
            } else {
                pictures = getSheetPictrues03((HSSFSheet) book.getSheetAt(i), (HSSFWorkbook) book);
            }

            try {
                List<T> tList = importSheet(book.getSheetAt(i), pojoClass);
                result.addAll(tList);
            } catch (Exception e) {
                e.printStackTrace(); // 跳过不识别的sheet，以及识别错误的字段, 读取单个sheet、读取错误字段异常
            }

        }

        return result;
    }


    // update-begin--Author:xuelin  Date:20171205 for：TASK #2098 【excel问题】 Online 一对多导入失败--------------------
    // result isn't used, but returned list is used, and it will be added into the result
    // import for every sheet in the Excel
    private <T> List<T> importSheet(Sheet sheet, Class<T> pojoClass) throws Exception {
        List<T> collection = new ArrayList<>();

        if (Map.class.equals(pojoClass)) {
            throw new Exception("not support Map format"); // this case is removed
        }

        Field[] fields = ReflectionUtil.getClassFields(pojoClass);
        // excelParams include all the information on the corresponding field
        getAllExcelField(fields, excelParams, excelCollection, pojoClass);
//        ignoreHeaderHandler(excelParams, params); // some group fields are chosen not to be imported
        Iterator<Row> rows = sheet.rowIterator();

        titleMap = getTitleMap(sheet, rows); // 读取表头信息

        //update-begin-author:liusq date:20220310 for:[issues/I4PU45]@excel里面新增属性fixedIndex
        Set<String> keys = excelParams.keySet();
        for (String key : keys) {
            if (key.startsWith("FIXED_")) {
                String[] arr = key.split("_");
                titleMap.put(Integer.parseInt(arr[1]), key); // put 1 , FIXED_1_Title
            }
        }
        //update-end-author:liusq date:20220310 for:[issues/I4PU45]@excel里面新增属性fixedIndex

        Row row = null;
        //跳过表头和标题行
        for (int j = 0; j < params.getTitleRows() + params.getHeadRows(); j++) {
            row = rows.next();
        }
        T object;

        while (rows.hasNext() && (row == null || sheet.getLastRowNum() - row.getRowNum() > params.getLastOfInvalidRow())) {
            row = rows.next();
            object = createObject(pojoClass);
            int finishedNum = importRow(row, object);
            if (finishedNum > 0 || collection.size() == 0) {
                collection.add(object);
            }
            for (ExcelCollectionParams param : excelCollection) { // one-to-many list construction branch
                addListContinue(collection.get(collection.size() - 1), param, row);
            }

        }
        return collection;
    }


    // firstCellNum, lastCellNum, titleMap, pojoClass, object, pictures, ImportParam params, Row, row
    private int importRow(Row row, Object object) throws Exception {
        Set<Integer> columnIndexSet = titleMap.keySet();
        Integer maxColumnIndex = Collections.max(columnIndexSet);
        Integer minColumnIndex = Collections.min(columnIndexSet);
        int firstCellNum = row.getFirstCellNum();
        if (firstCellNum > minColumnIndex) {
            firstCellNum = minColumnIndex;
        }
        int lastCellNum = row.getLastCellNum();
        if (lastCellNum < maxColumnIndex + 1) {
            lastCellNum = maxColumnIndex + 1;
        }
        ExcelImportEntity rowEntity = excelParams.get(ROW_FIELD);
        if (rowEntity != null) {
            int rowNum = row.getRowNum();
            rowEntity.getMethod().invoke(object, rowNum + 1);  // set row number into the created object
        }
        int finishedNum = 0;
        String picId;
        for (int i = firstCellNum; i < lastCellNum; i++) {
            Cell cell = row.getCell(i);
            String titleString = titleMap.get(i);
            if (!excelParams.containsKey(titleString)) continue;
            if (excelParams.get(titleString) != null && excelParams.get(titleString).getType() == 2) { // 先处理图片格式的情况
                picId = row.getRowNum() + "_" + i;
                saveImage(object, picId, excelParams, titleString, pictures, params);
            } else if (saveFieldValue(object, cell, excelParams, titleString)) { // the general case, to set value from Excel into object
                finishedNum++;
            }
        }
        return finishedNum;
    }


    /**
     * 保存字段值(获取值,校验值,追加错误信息)
     */
    private boolean saveFieldValue(Object object, Cell cell, Map<String, ExcelImportEntity> excelParams, String titleString) throws Exception {
        Object value = cellValueServer.getValue(object, cell, excelParams.get(titleString));
        if (value instanceof String && ((String) value).trim().equals("")) return false;
        excelParams.get(titleString).getMethod().invoke(object, value);
        return value != null;
    }


    /**
     * 获取需要导入的全部字段
     */
    public static void getAllExcelField(Field[] fields, Map<String, ExcelImportEntity> excelParams,
                                        List<ExcelCollectionParams> excelCollection, Class<?> pojoClass) throws Exception {
        for (Field field : fields) {
            if (field.getAnnotation(ExcelCollection.class) != null && ReflectionUtil.isCollection(field.getType())) {
                LOGGER.debug("read collection field : {} , of tyep : {}", field.getName(), field.getType());
                // 集合对象设置属性
                ExcelCollectionParams collection = new ExcelCollectionParams(field);
                excelCollection.add(collection);
                getAllExcelFieldForList(collection, ReflectionUtil.getClassFields(collection.getType()),
                        collection.getType()); // no third level is considered
            } else if (field.getAnnotation(ExcelField.class) != null && ReflectionUtil.isJavaClass(field)) {
                LOGGER.debug("read java class field : {} , of tyep : {}", field.getName(), field.getType());
                addEntityToMap(field, pojoClass, excelParams); // read @Excel annotation

                // if not basic class, consider it as a user-defined class,
                // fields of which should be basic javaClass to recursively obtain field information
            } else if (field.getAnnotation(ExcelEntity.class) != null) {
                LOGGER.debug("read else field : {} , of tyep : {}", field.getName(), field.getType());

                getAllExcelField(ReflectionUtil.getClassFields(field.getType()), excelParams, excelCollection, field.getType());
            }
        }
    }

    /**
     * 把这个注解解析放到类型对象中
     *
     * @throws Exception
     */
    static void addEntityToMap(Field field, Class<?> pojoClass,
                               Map<String, ExcelImportServer.ExcelImportEntity> temp) throws Exception {
        ExcelImportServer.ExcelImportEntity excelEntity = ExcelImportEntity.generateEntity(field, pojoClass);
        addEntityToMap(excelEntity, null, temp);
    }

    static void addEntityToMap(ExcelImportServer.ExcelImportEntity excelEntity, String collectionName, Map<String, ExcelImportServer.ExcelImportEntity> temp) {
        StringBuilder prefix = new StringBuilder("");
        if (StringUtils.isNotEmpty(collectionName)) prefix.append(collectionName).append("_");
        if (StringUtils.isNotEmpty(excelEntity.getGroupName())) prefix.append(excelEntity.getGroupName()).append("_");
        if (excelEntity.getFixedIndex() != -1) {
            prefix.append("FIXED_").append(excelEntity.getFixedIndex());
            temp.put(prefix + excelEntity.getName(), excelEntity);
            return;
        }
        temp.put(prefix + excelEntity.getName(), excelEntity);
    }


    public static void getAllExcelFieldForList(ExcelImportServer.ExcelCollectionParams collection, Field[] fields, Class<?> pojoClass
    ) throws Exception {
        Map<String, ExcelImportServer.ExcelImportEntity> temp = collection.getExcelParams();
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (field.getAnnotation(ExcelField.class) != null &&
                    ReflectionUtil.isJavaClass(field)) {
                ExcelImportServer.ExcelImportEntity excelEntity = ExcelImportEntity.generateEntity(field, pojoClass);
                addEntityToMap(excelEntity, collection.getExcelName(), temp);
            } else if (field.getAnnotation(ExcelCollection.class) != null &&
                    ReflectionUtil.isCollection(field.getType())) {
                // collection should also contain sub-collections
                // create new collection here , put it as a sub-collection
                ExcelImportServer.ExcelCollectionParams subCollection = new ExcelImportServer.ExcelCollectionParams(field);
                subCollection.setExcelName(collection.getExcelName() + "_" + subCollection.getExcelName());
                collection.getSubCollections().add(subCollection);
                // collection case, need recursively construct sub colletion
                getAllExcelFieldForList(subCollection, ReflectionUtil.getClassFields(subCollection.getType()), subCollection.getType());
            }
        }
    }

    //endregion

    /***
     * 向List里面继续添加元素
     *
     */
    private void addListContinue(Object object, ExcelCollectionParams param, Row row) throws Exception {
        List collection = (List) ReflectionUtil.getGetMethod(param.getName(), object.getClass()).invoke(object, new Object[]{});
        if (collection == null)
            throw new RuntimeException("collection should have been initialized with createObject before");
        Object entity = createObject(param.getType()); // all ExcelCollection-annotated field are initialized

        boolean isUsed = false;// 是否需要加上这个对象
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) { // similar with importRow method todo
            Cell cell = row.getCell(i);
            String titleString = (String) titleMap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                if (param.getExcelParams().get(titleString).getType() == 2) {
                    String picId = row.getRowNum() + "_" + i;
                    saveImage(object, picId, param.getExcelParams(), titleString, pictures, params);
                } else if (cell == null) {
                    continue;
                } else {
                    saveFieldValue(entity, cell, param.getExcelParams(), titleString);
                }
                isUsed = true;
            }
        }

        // if isUsed = false and collection is empty create a empty one for recursive construct
        if (isUsed || collection.size() == 0) {
            ExcelImportEntity rowEntity = param.getExcelParams().get(param.getExcelName() + ROW_FIElD);
            if (rowEntity != null) {
                int rowNum = row.getRowNum();
                rowEntity.getMethod().invoke(entity, rowNum + 1);  // set row number into the created object
            }
            collection.add(entity);
        }

        List<ExcelCollectionParams> subCollections = param.getSubCollections();
        if (subCollections != null && subCollections.size() > 0) {
            for (ExcelCollectionParams subCollection : subCollections) {
                // call the method recursively
                addListContinue(collection.get(collection.size() - 1), subCollection, row);
            }
        }
    }


    private Map<Integer, String> getTitleMap(Sheet sheet, Iterator<Row> rows) throws Exception {
        //update_begin-author:taoyan date:2020622 for：当文件行数小于代码里设置的TitleRows时headRow一直为空就会出现死循环
        //找到首行表头，每个sheet都必须至少有一行表头
        Row headRow = null;
        int headBegin = params.getTitleRows();
        int allRowNum = sheet.getPhysicalNumberOfRows();
        while (headRow == null && headBegin < allRowNum) {
            headRow = sheet.getRow(headBegin++);
        }
        if (headRow == null) {
            throw new Exception("表头为空");
        }
        StringBuilder debugInfo1 = new StringBuilder();
        HashMap<Integer, Integer> previousMergedRow = null;
        Map<Integer, String> titlemap = new HashMap<Integer, String>();
        for (int j = headBegin - 1; ; j++) { // j < headBegin + params.getHeadRows() - 1
            Row currentRow = sheet.getRow(j);

            // determine whether the next-row read can be on going , which is whether it has reached data region
            if (previousMergedRow != null) { // first row shouldn't skipped
                Iterator<Cell> cellIterator2 = currentRow.cellIterator();
                boolean shouldJumpOut = false;
                boolean allOutOfRegion = true;
                while (cellIterator2.hasNext()) {
                    Cell cell = cellIterator2.next();
                    boolean isInPreviousMergeRegion = isInPreviousMergeRegion(cell.getColumnIndex(), previousMergedRow);
                    if (!isInPreviousMergeRegion) allOutOfRegion = false;
                    if (!isInPreviousMergeRegion && StringUtils.isNotEmpty(cell.getStringCellValue())) {

                        shouldJumpOut = true;
                        break;
                    }
                }
                if (shouldJumpOut || allOutOfRegion) break;
            }

            HashMap<Integer, Integer> currentMergedRow = new HashMap<>();
            Iterator<Cell> cellIterator1 = currentRow.cellIterator();
            String previousValue = null;
            Integer previousMergeRegionBegin = null;
            while (cellIterator1.hasNext()) {
                Cell cell = cellIterator1.next();
                String value = cell.getStringCellValue();
                debugInfo1.append(" ").append(value);
                boolean isInPreviousMergeRegion = true;
                if (previousMergedRow != null)// determined by current row and previousMergeRegionBegin
                    isInPreviousMergeRegion = isInPreviousMergeRegion(cell.getColumnIndex(), previousMergedRow);
                if (StringUtils.isNotEmpty(value)) {
                    String current = titlemap.get(cell.getColumnIndex());
                    if (StringUtils.isNotEmpty(current)) {
                        titlemap.put(cell.getColumnIndex(), current + "_" + value);//加入表头列表
                        previousValue = current + "_" + value;
                    } else {
                        titlemap.put(cell.getColumnIndex(), value);//加入表头列表
                        previousValue = value;
                    }
                    previousMergeRegionBegin = null;

                    //    one(parent)-to-one(child) case, that is one parent has only one child, this case is excluded.
                    // Because to determine it has child to be considered as a mergedRegion needs next row information,
                    // causing this case clash with case where next row is all excel data not row head,
                    // so it is difficult to determine how many number of rows to be considered as head row.
                    //    In another way, to introduce this case, you should set params.getHeadRows() in advance,
                    // and not using above loop to determine number of head rows.
//                    if (!isInPreviousMergeRegion) currentMergedRow.put(cell.getColumnIndex(), cell.getColumnIndex());

                } else if (StringUtils.isNotEmpty(previousValue) && isInPreviousMergeRegion) { // and also should be whithin merge region of previous row todo
                    if (previousMergeRegionBegin == null) previousMergeRegionBegin = cell.getColumnIndex() - 1;
                    currentMergedRow.put(previousMergeRegionBegin, cell.getColumnIndex());
                    titlemap.put(cell.getColumnIndex(), previousValue);//加入表头列表
                } else if (!isInPreviousMergeRegion) {
                    // case1, empty value and also out of merge region, previous value should be set empty, then skip
                    previousValue = null;
                } else if (StringUtils.isEmpty(previousValue)) {
                    // case2, empty value, but whithin merge region, empty previous value, throw runtime error , read head row eception,
                    // the current cell shouldn't be empty
                    throw new Exception("表头格式错误， " + cell.getRowIndex() + "行 " + cell.getColumnIndex() + "列 " + "不应为空");
                }
            }
            previousMergedRow = currentMergedRow;
        }

        debugInfo1.append("\n");
        LOGGER.debug(debugInfo1.toString());
        return titlemap;
    }

    private static boolean isInPreviousMergeRegion(Integer inputColunm, HashMap<Integer, Integer> mergedRow) {
        for (Integer beginColumn : mergedRow.keySet()) {
            Integer endColumn = mergedRow.get(beginColumn);
            if (inputColunm <= endColumn && inputColunm >= beginColumn)
                return true;
        }
        return false;
    }


    public static class ExcelImportEntity {

        /**
         * 对应name
         */
        protected String name;
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
         * 表头组名称
         */
        private String groupName;

        /**
         * set/get方法
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

//        public String getFormulaExpr() {
//            return formulaExpr;
//        }

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

        public String getGroupName() {
            return groupName;
        }

        public void setGroupName(String groupName) {
            this.groupName = groupName;
        }

        public Integer getFixedIndex() {
            return fixedIndex;
        }

        public void setFixedIndex(Integer fixedIndex) {
            this.fixedIndex = fixedIndex;
        }

        /**
         * 对应 Collection NAME
         */
        private String collectionName;
        /**
         * 保存图片的地址 当saveType设置为3/4时，此值可以设置为：local,minio,alioss
         */
        private String saveUrl;
        /**
         * 保存图片的类型,1是文件_old,2是数据库字节,3文件地址_new,4网络地址
         */
        private int saveType;

        private Class<?> classType2;

        private List<ExcelImportEntity> list;

//        public String getClassType() {
//            return classType;
//        }

        public Class<?> getClassType2() {
            return classType2;
        }

//        public String getCollectionName() {
//            return collectionName;
//        }

        public List<ExcelImportEntity> getList() {
            return list;
        }

        public int getSaveType() {
            return saveType;
        }

        public String getSaveUrl() {
            return saveUrl;
        }

//        public void setClassType(String classType) {
//            this.classType = classType;
//        }

        public void setClassType2(Class<?> classType) {
            this.classType2 = classType;
        }

//        public void setCollectionName(String collectionName) {
//            this.collectionName = collectionName;
//        }

        public void setList(List<ExcelImportEntity> list) {
            this.list = list;
        }

        public void setSaveType(int saveType) {
            this.saveType = saveType;
        }

        public void setSaveUrl(String saveUrl) {
            this.saveUrl = saveUrl;
        }



        public static ExcelImportServer.ExcelImportEntity generateEntity(Field field,
                                                                         Class<?> pojoClass) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            ExcelImportServer.ExcelImportEntity excelEntity = new ExcelImportServer.ExcelImportEntity();
            excelEntity.setType(excelField.type());
            excelEntity.setSaveUrl(excelField.savePath());
            excelEntity.setSaveType(excelField.imageType());
            excelEntity.setReplace(excelField.replace());
            excelEntity.setDatabaseFormat(excelField.databaseFormat());

            excelEntity.setNumFormat(excelField.numFormat());
            excelEntity.setGroupName(excelField.groupName());
            //update-begin-author:liusq date:20220310 for:[issues/I4PU45]@excel里面新增属性fixedIndex
            excelEntity.setFixedIndex(excelField.fixedIndex());
            //update-end-author:liusq date:20220310 for:[issues/I4PU45]@excel里面新增属性fixedIndex

            //update-begin-author:taoYan date:20180202 for:TASK #2067 【bug excel 问题】excel导入字典文本翻译问题
            excelEntity.setMultiReplace(excelField.multiReplace());
            //update-end-author:taoYan date:20180202 for:TASK #2067 【bug excel 问题】excel导入字典文本翻译问题


            getExcelField(field, excelEntity, excelField, pojoClass);

            return excelEntity;
        }

        private static void getExcelField(Field field, ExcelImportServer.ExcelImportEntity excelEntity, ExcelField excelField, Class<?> pojoClass) {
            excelEntity.setName(excelField.name());
            String fieldname = field.getName();
            excelEntity.setClassType2(field.getType());
            //update-begin-author:taoyan for:TASK #2798 【例子】导入扩展方法，支持自定义导入字段转换规则
            try {
                excelEntity.setMethod(getSetMethod(fieldname, pojoClass, field.getType(), excelField.importConvert()));
            } catch (Exception e) {
                LOGGER.error(" method not found for the field  {}  excel entity", field.getName());
                throw new RuntimeException(e);
            }
            //update-end-author:taoyan for:TASK #2798 【例子】导入扩展方法，支持自定义导入字段转换规则
            if (StringUtils.isNotEmpty(excelField.importFormat())) {
                excelEntity.setFormat(excelField.importFormat());
            } else {
                excelEntity.setFormat(excelField.format());
            }
            excelEntity.setImportFormats(excelField.importFormats());
        }
    }

    public static class ExcelCollectionParams {

        /**
         * 集合对应的名称
         */
        private String name;
        /**
         * Excel 列名称
         */
        private String excelName;
        /**
         * 实体对象
         */
        private Class<?> type;
        /**
         * 这个list下面的参数集合实体对象
         */
        private final Map<String, ExcelImportEntity> excelParams = new HashMap<>();
        ;

        private final List<ExcelCollectionParams> subCollections = new ArrayList<>();

//        public ExcelCollectionParams() {
//        }

        public ExcelCollectionParams(Field field) {
            this.name = field.getName();
            ParameterizedType pt = (ParameterizedType) field.getGenericType(); // object type of the list in <>
            this.type = (Class<?>) pt.getActualTypeArguments()[0];
            this.excelName = field.getAnnotation(ExcelCollection.class).name();
        }

        public List<ExcelCollectionParams> getSubCollections() {
            return subCollections;
        }

        public Map<String, ExcelImportEntity> getExcelParams() {
            return excelParams;
        }

        public String getName() {
            return name;
        }

        public Class<?> getType() {
            return type;
        }

//        public void setExcelParams(Map<String, ExcelImportEntity> excelParams) {
//            this.excelParams = excelParams;
//        }

        public void setName(String name) {
            this.name = name;
        }

        public void setType(Class<?> type) {
            this.type = type;
        }

//        public void setSubCollections(List<ExcelCollectionParams> subCollections) {
//            this.subCollections = subCollections;
//        }


        public String getExcelName() {
            return excelName;
        }

        public void setExcelName(String excelName) {
            this.excelName = excelName;
        }
    }
}
