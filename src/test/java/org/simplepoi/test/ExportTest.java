package org.simplepoi.test;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.simplepoi.excel.ExcelExportUtil;
import org.simplepoi.excel.ExportParams;
import org.simplepoi.excel.constant.ExcelType;
import org.simplepoi.test.entity.FormulaEntity;
import org.simplepoi.test.entity.WidthTestEntity1;
import org.simplepoi.test.entity.WidthTestEntity2;
import org.simplepoi.test.tokenization.GenericTokenParser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExportTest {
    // output files to desktop
    //  new File("C:\\Users\\Administrator\\Desktop\\日期测试.xlsx");
    @Test
    public void test1() {


    }

    // formula export test, use tokenizer to convert expression into corresponding cell or array of cells

    // * IF(SUM(#{contractMoney1} + #{contractMoney2})>0,1,2 )  SUM(B2)
    // * annotation design @Excel(type = 3, formulaExpr = "SUM(#{contractMoney}))

    // For one-to-many case, obtain sum of a field of multiple rows (the "many" part)
    // * SUM(#{sum_contractMoney}) -> SUM(A1:A9) , For all many part
    // * subtype consideration ?  for onw-to-many case, the "one" set its value to sum of the "many".
    // generally for sum function, used like A1:A2
    // 1, determine which field is chosen as subtype in list, use sumType in annotation @Excel,
    // there is only one sumType for a pojo class, generally(must for present) be Integer, but also can be String
    // 2, determine which subtype is used for the sum function #{sum_1_consideration} , presently
    // constrain only one Integer field to indicate the subtype


    // ExcelExportBase.setCellWith() 有问题 组的情况 基于 ExcelParams 设置列宽时 考虑错误
    @Test
    public void widthTest() { // 宽度设置错位，对后一列生效
        ArrayList<Map<String, Object>> listMap = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("title", new ExportParams(null, "宽度测试1", ExcelType.XSSF));
        List<WidthTestEntity1> earningStatData = new ArrayList<>();
        map.put("data", earningStatData);
        map.put("entity", WidthTestEntity1.class);
        listMap.add(map);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("title", new ExportParams(null, "宽度测试2", ExcelType.XSSF));
        List<WidthTestEntity2> earningStatData2 = new ArrayList<>();
        map2.put("data", earningStatData2);
        map2.put("entity", WidthTestEntity2.class);
        listMap.add(map2);

        exportToFile(listMap, "宽度测试");  // 文件导出
    }

    // for some situations, width set in the @Excel annotation doesn't take effect, worth investigation

    private void exportToFile(ArrayList<Map<String, Object>> listMap, String fileName) {
        Workbook workbook = ExcelExportUtil.exportExcel(listMap, ExcelType.XSSF);
        File file = new File("C:\\Users\\Administrator\\Desktop\\" + fileName + new SimpleDateFormat("yyyy-MM-dd").format(new Date()) + ".xlsx");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    @Test
    public void formulaExportTest() {
        List<FormulaEntity> formulaEntities = new ArrayList<>();
        formulaEntities.add(new FormulaEntity(new BigDecimal(200), new BigDecimal(300), new BigDecimal(500)));
        formulaEntities.add(new FormulaEntity(new BigDecimal(500), new BigDecimal(300), new BigDecimal(800)));

        ArrayList<Map<String, Object>> listMap = new ArrayList<>();
        Map<String, Object> map3 = new HashMap<>(); // 该情况 偏移叫大，前对后面的两列生效
        map3.put("title", new ExportParams(null, "公式导出测试", ExcelType.XSSF));
        map3.put("data", formulaEntities);
        map3.put("entity", FormulaEntity.class);
        listMap.add(map3);
        exportToFile(listMap, "公式导出测试");  // 文件导出

        System.out.println("ok");

    }

    @Test
    public void tokenParseTest(){
        Properties testProp = new Properties();
        testProp.setProperty("aaa1","1234"); // return null, the previous value
        testProp.setProperty("aaa2","2221234");
        String openToken = "${",closeToken = "}";
      //  VariableTokenHandler handler = new VariableTokenHandler(testProp,openToken,closeToken);
        GenericTokenParser parser = new GenericTokenParser(openToken, closeToken, testProp);
        String parseResult = parser.parse("${aaa1}}${aaa2}");
        System.out.println(parseResult);
        System.out.println("ok");
    }

}
