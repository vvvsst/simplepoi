package org.simplepoi.test.treeimporttest.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelField;

import java.util.List;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class Student {
    /**姓名*/
    @ExcelField(name = "姓名", width = 15)
    private String name; //姓名
    /**年龄*/
    @ExcelField(name = "年龄", width = 15)
    private String age; // 年龄
    /**性別*/
    @ExcelField(name = "性別", width = 15)
    private String sex; // 性別

    @ExcelField(name =ROW_FIElD, type = 5)
    private Integer lineNum ;

    @ExcelCollection(name = "sub学生1")
    private List<Student2> studentList1;
    @ExcelCollection(name = "sub学生2")
    private List<Student2> studentList2;

    public Student(String name, String age, String sex) {
        this.name = name;
        this.age = age;
        this.sex = sex;
    }
}
