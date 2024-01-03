package org.simplepoi.test.treeimporttest.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.simplepoi.excel.annotation.ExcelField;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class Student2 {
    /**姓名*/
    @ExcelField(name = "姓名2", width = 15)
    private String name; //姓名
    /**年龄*/
    @ExcelField(name = "年龄2", width = 15)
    private String age; // 年龄
    /**性別*/
    @ExcelField(name = "性別2", width = 15)
    private String sex; // 性別

    @ExcelField(name = ROW_FIElD, type = 5)
    private Integer lineNum ;


    public Student2(String name, String age, String sex) {
        this.name = name;
        this.age = age;
        this.sex = sex;
    }
}
