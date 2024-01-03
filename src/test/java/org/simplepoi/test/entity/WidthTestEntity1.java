package org.simplepoi.test.entity;

import lombok.Data;
import org.simplepoi.excel.annotation.ExcelField;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class WidthTestEntity1 {

    @ExcelField(name = "日期", format = "yyyy年M月", width = 30) // yyyy-MM-dd 只取年月即可 yyyy-MM-
    private LocalDate theDate;



    @ExcelField(name = "金额1(元)",groupName = "合计", width = 40, height = 15, orderNum = "2")  // 需要多个合同 合并
    private BigDecimal contractMoney;

    @ExcelField(name = "金额2(元)",groupName = "合计", width = 30, height = 15, orderNum = "3")
    private BigDecimal invoicedMoney;

    @ExcelField(name = "金额3(元)",groupName = "合计", width = 20, height = 15, orderNum = "3")
    private BigDecimal invoicedMoney22;

    @ExcelField(name = "金额4(元)",groupName = "合计", width = 10, height = 15, orderNum = "3")
    private BigDecimal invoicedMoney2;

    public WidthTestEntity1() {
    }
}
