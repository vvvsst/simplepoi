package org.simplepoi.functest.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelField;
import org.simplepoi.functest.FuncTestUtils;

import java.util.ArrayList;
import java.util.List;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class Student {
    /**姓名*/
    @ExcelField(name = "姓名", width = 15,groupName = "基本信息")
    private String name; //姓名
    /**年龄*/
    @ExcelField(name = "年龄", width = 15,groupName = "基本信息")
    private Integer age = FuncTestUtils.randomAge(); // 年龄

    /**性別*/
    @ExcelField(name = "性别", width = 15,type = 1,replace ={"男_1","女_0"},groupName = "基本信息")
    private Integer gender = FuncTestUtils.randomZeroOrOne();

    @ExcelCollection(name = "成绩")
    private List< Grade> gradeList = new ArrayList<>();

    @ExcelField(name = ROW_FIElD, type = 5)
    private Integer lineNum ;

    public Student(String name) {
        this.name = name;
    }

    public Student addGrade(Grade grade){
        this.gradeList.add(grade);
        return this;
    }

}
