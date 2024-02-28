package org.simplepoi.test.entity;

import lombok.Data;
import org.simplepoi.excel.annotation.ExcelField;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class WidthTestEntity3 {

    @ExcelField(name = "月份", format = "yyyy年M月", width = 20) // yyyy-MM-dd 只取年月即可 yyyy-MM-
    private LocalDate theDate;

    //region 部门信息
    // 以下这两个字段 一定会导出
    @ExcelField(name = "合同单位", width = 20, height = 15, orderNum = "0")
    private String payerComp;

    @ExcelField(name = "收入单位", width = 20, height = 15, orderNum = "1")
    private String payeeComp;

    private Long payeeCompId; // a condition used for merging data

    private Long payerCompId; // a condition used for merging data

    private int hasAllDiffer = 0; // used for mark whether the data of differ1 and differ2 all has been added, 0(none), 1(one part) , 2(finished)
    //endregion


    //region 总额
    // 代营 与 划小 之和

    // 即所关联的所有合同的合同金额之和
    @ExcelField(name = "应开票总金额(元)",groupName = "合计", width = 20, height = 15, orderNum = "2")  // 需要多个合同 合并
    private BigDecimal contractMoney;

    // 取值为 与  应收总金额 相等
    @ExcelField(name = "已开票总金额(元)",groupName = "合计", width = 20, height = 15, orderNum = "3")
    private BigDecimal invoicedMoney;

    // 这些收入记录所记录的 合同中所关联的 收入记录之和
    // @Excel(name = "已开票总金额", width = 20, height = 15, orderNum = "3")
    private BigDecimal earningMoneyBound; // 需要多个合同 合并

    // 取值为所关联的 合同 中 未开票总金额，即 应开票总金额 contractMoney - 合同已绑定的总金额 earningMoneyBound
    @ExcelField(name = "未开票总金额(元)",groupName = "合计", width = 20, height = 15, orderNum = "4") // 需要多个合同 合并
    private BigDecimal notInvoicedMoney;

    // 这些收入记录所记录的 含税价之和 taxMoneyPrice 而非合同
    @ExcelField(name = "应收总金额(元)",groupName = "合计", width = 20, height = 15, orderNum = "5")
    private BigDecimal taxMoneyPrice;

    // 应收总金额 不含税 即 价款字段
    @ExcelField(name = "应收总金额(不含税)(元)",groupName = "合计", width = 20, height = 15, orderNum = "6")
    private BigDecimal priceMoney;

    // 回款总金额 ， 根据未回款总金额计算而来
    @ExcelField(name = "回款总金额(元)",groupName = "合计", width = 20, height = 15, orderNum = "7")
    private BigDecimal finishedMoney;

    // 未回款总金额，即 pending money
    @ExcelField(name = "未回款总金额(元)",groupName = "合计", width = 20, height = 15, orderNum = "8")
    private BigDecimal pendingMoney;

    // 计算方法为 应开票总金额 contractMoney  - 已开票总金额 invoicedMoney
    @ExcelField(name = "差额合计(元)",groupName = "合计", width = 20, height = 15, orderNum = "9")
    private BigDecimal differenceMoney;

    //endregionW

    //region  划小部分

    @ExcelField(name = "划小应开票总金额(元)",groupName = "划小", width = 50, height = 15, orderNum = "10")  // 需要多个合同 合并
    private BigDecimal contractMoneyDiffer1 = new BigDecimal(0);

    @ExcelField(name = "划小已开票总金额(元)",groupName = "划小", width = 20, height = 15, orderNum = "11")
    private BigDecimal invoicedMoneyDiffer1 = new BigDecimal(0);

    private BigDecimal earningMoneyBoundDiffer1 = new BigDecimal(0); // 需要多个合同 合并

    @ExcelField(name = "划小未开票总金额(元)",groupName = "划小", width = 20, height = 15, orderNum = "12") // 需要多个合同 合并
    private BigDecimal notInvoicedMoneyDiffer1 = new BigDecimal(0);

    @ExcelField(name = "划小应收总金额(元)",groupName = "划小", width = 20, height = 15, orderNum = "13")
    private BigDecimal taxMoneyPriceDiffer1 = new BigDecimal(0);

    @ExcelField(name = "划小应收总金额(不含税)(元)",groupName = "划小", width = 20, height = 15, orderNum = "14")
    private BigDecimal priceMoneyDiffer1 = new BigDecimal(0);

    @ExcelField(name = "划小回款总金额(元)",groupName = "划小", width = 20, height = 15, orderNum = "15")
    private BigDecimal finishedMoneyDiffer1 = new BigDecimal(0);

    @ExcelField(name = "划小未回款总金额(元)",groupName = "划小", width = 20, height = 15, orderNum = "16")
    private BigDecimal pendingMoneyDiffer1 = new BigDecimal(0);

    @ExcelField(name = "划小差额合计(元)",groupName = "划小", width = 20, height = 15, orderNum = "17")
    private BigDecimal differenceMoneyDiffer1 = new BigDecimal(0);

    private String contractIdsDiffer1 = "";
    //endregion


    //region  代营部分

    @ExcelField(name = "代营应开票总金额(元)",groupName = "代营", width = 20, height = 15, orderNum = "18")  // 需要多个合同 合并
    private BigDecimal contractMoneyDiffer2 = new BigDecimal(0);

    @ExcelField(name = "代营已开票总金额(元)",groupName = "代营", width = 20, height = 15, orderNum = "19")
    private BigDecimal invoicedMoneyDiffer2 = new BigDecimal(0);

    private BigDecimal earningMoneyBoundDiffer2 = new BigDecimal(0); // 需要多个合同 合并

    @ExcelField(name = "代营未开票总金额(元)",groupName = "代营", width = 20, height = 15, orderNum = "20") // 需要多个合同 合并
    private BigDecimal notInvoicedMoneyDiffer2 = new BigDecimal(0);

    @ExcelField(name = "代营应收总金额(元)",groupName = "代营", width = 20, height = 15, orderNum = "21")
    private BigDecimal taxMoneyPriceDiffer2 = new BigDecimal(0);

    @ExcelField(name = "代营应收总金额(不含税)(元)",groupName = "代营", width = 30, height = 15, orderNum = "22")
    private BigDecimal priceMoneyDiffer2 = new BigDecimal(0);

    @ExcelField(name = "代营回款总金额(元)",groupName = "代营", width = 20, height = 15, orderNum = "23")
    private BigDecimal finishedMoneyDiffer2 = new BigDecimal(0);

    @ExcelField(name = "代营未回款总金额(元)",groupName = "代营", width = 20, height = 15, orderNum = "24")
    private BigDecimal pendingMoneyDiffer2 = new BigDecimal(0);

    @ExcelField(name = "代营差额合计(元)",groupName = "代营", width = 20, height = 15, orderNum = "25")
    private BigDecimal differenceMoneyDiffer2 = new BigDecimal(0);

    private String contractIdsDiffer2 = "";

    //endregion

    public WidthTestEntity3() {
    }
}
