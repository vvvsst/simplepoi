package org.simplepoi.functest.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.simplepoi.excel.annotation.ExcelField;
import org.simplepoi.functest.FuncTestUtils;

import java.math.BigDecimal;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class Grade {
    /**姓名*/
    @ExcelField(name = "科目名称", width = 15)
    private String name; //姓名

    @ExcelField(name = "分数1",type = 4, width = 10 ,groupName = "评分")
    private BigDecimal score1 = FuncTestUtils.randomBigDecimal();

    @ExcelField(name = "分数2",type = 4, width = 10 ,groupName = "评分")
    private BigDecimal score2 = FuncTestUtils.randomBigDecimal();

    @ExcelField(name = "均分",type = 3, formulaExpr = "(${COL_score1}${ROW_LIST} + ${COL_score2}${ROW_LIST})/2 " ,width = 10 ,groupName = "评分")
    private BigDecimal score3;

    @ExcelField(name = ROW_FIElD, type = 5)
    private Integer lineNum ;


    public Grade(String name) {
        this.name = name;
    }
}
