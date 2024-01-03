package org.simplepoi.test.entity;


import lombok.Data;
import org.simplepoi.excel.annotation.ExcelField;

import java.math.BigDecimal;
import java.util.Date;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@Data
public class BasicTestEntity {

    @ExcelField(name = "文本")
    private String normalStr;

    @ExcelField(name = "类型", replace = {"aaaa_1","bbb_20","已审核_50"})
    private Integer type;

    @ExcelField(name = "文本类型测试")
    private String dataFormatTest;

    @ExcelField(name = "日期Date")
    private Date normalDate;

    @ExcelField(name = "日期LocalDate")
    private Date normalLocalDate;

    @ExcelField(name = "整数")
    private Integer normalInt;

    @ExcelField(name = "长整数")
    private Long normalLong;

    @ExcelField(name = "金额", type = 4)
    private BigDecimal normalDecimal;

    @ExcelField(name = ROW_FIElD, type = 5)  // do not export , and import line number of Excel
    private Integer rowNum;

}
