package org.simplepoi.test.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.simplepoi.excel.annotation.ExcelField;

import java.time.LocalDate;
import java.util.Date;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class DateEntity {

    @ExcelField(name = "日期1", width = 15,format = "yyyy-MM-dd", orderNum = "1")
    private Date date1;

    @ExcelField(name = "日期2", width = 15, format = "yyyy-MM-dd", orderNum = "2")
    private LocalDate localDate2;

    @ExcelField(name = "日期3", width = 15 , orderNum = "3")
    private Date date3;

    @ExcelField(name = "日期4", width = 15 , orderNum = "4")
    private LocalDate localDate4;

    @ExcelField(name = "测试5", width = 15 , orderNum = "4")
    private String testStr = "aaaaaaaaaaccc";

    @ExcelField(name = "合同状态",  width = 20)
    private String contractStatus;

    @ExcelField(name =ROW_FIElD, type = 5)
    private Integer lineNum ;

    public DateEntity(Date date1, LocalDate localDate2, Date date3, LocalDate localDate4) {
        this.date1 = date1;
        this.localDate2 = localDate2;
        this.date3 = date3;
        this.localDate4 = localDate4;
    }
}
