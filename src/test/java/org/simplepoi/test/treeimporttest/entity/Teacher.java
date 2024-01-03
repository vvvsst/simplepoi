package org.simplepoi.test.treeimporttest.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelField;

import java.math.BigDecimal;
import java.util.List;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class Teacher {
    /**
     * 教师名称
     */
    @ExcelField(name = "教师名称", width = 15,type = 1)
    private String name;

    @ExcelField(name = "性质2", width = 15,type = 4)
    private BigDecimal property2;

    @ExcelField(name = "性质3", width = 15,type = 4)
    private BigDecimal property3;

    @ExcelField(name = "函数测试", type = 3 )
    private String property4 = "SUM(B6:C6)";

    @ExcelField(name = "图片", type = 2, imageType = 4)
    private String figUrl = "https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png";
    @ExcelField(name = ROW_FIElD, type = 5)
    private Integer lineNum ;
    /**
     * 学生
     */
    @ExcelCollection(name = "学生")
    private List<Student> studentList;
    @ExcelCollection(name = "学生2")
    private List<Student> studentList2;

    public Teacher(String name, String property2, String property3) {
        this.name = name;
        this.property2 = new BigDecimal(property2);
        this.property3 =new BigDecimal(property3);
    }
}
