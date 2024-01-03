package org.simplepoi.test.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.simplepoi.excel.annotation.ExcelField;

import java.math.BigDecimal;

@AllArgsConstructor
@Data
public class FormulaEntity {

    @ExcelField(name = "金额1", type = 4, width = 20,orderNum = "0")
    private BigDecimal testMoney1;

    @ExcelField(name = "金额2", type = 4, width = 20,orderNum = "1")
    private BigDecimal testMoney2; // 代加金额2

    @ExcelField(name = "金额3", type = 3, formulaExpr = "${testMoney1} + ${testMoney2} -20000", width = 20,orderNum = "2")
    private BigDecimal testMoney3;

}
