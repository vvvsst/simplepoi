package org.simplepoi.functest.entity;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.math3.stat.inference.TestUtils;
import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelField;
import org.simplepoi.excel.constant.PoiBaseConstants;
import org.simplepoi.functest.FuncTestUtils;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import static org.simplepoi.excel.constant.PoiBaseConstants.ROW_FIElD;

@NoArgsConstructor
@Data
public class Teacher {
    /**
     * 教师名称
     */
    @ExcelField(name = "教师名称", width = 15,type = 1,orderNum = "0")
    private String name;

    // 性别
    @ExcelField(name = "性别", width = 15,type = 1,replace ={"男_1","女_0"},orderNum = "1")
    private Integer gender  = FuncTestUtils.randomZeroOrOne();

    @ExcelField(name = "评分1", width = 10,type = 4,groupName = "评分",orderNum = "2")
    private BigDecimal score1 = FuncTestUtils.randomBigDecimal();

    @ExcelField(name = "评分2", width = 10,type = 4,groupName = "评分",orderNum = "3")
    private BigDecimal score2= FuncTestUtils.randomBigDecimal();;

    // PoiBaseConstants.VAR_COL
    @ExcelField(name = "总评分",type = 3, formulaExpr = "${COL_score1}${ROW_LIST} +${COL_score2}${ROW_LIST}", width = 10 ,groupName = "评分",orderNum = "4")
    private BigDecimal score3;

    @ExcelField(name =ROW_FIElD, type = 5) // 导入的时候获得行号
    private Integer lineNum ;
    /**
     * 学生
     */
    @ExcelCollection(name = "学生1班",orderNum = "5")
    private List<Student> studentList = new ArrayList<>();
    @ExcelCollection(name = "学生2班",orderNum = "6")
    private List<Student> studentList2 = new ArrayList<>();

    public Teacher(String name) {
        this.name = name;
    }
    public Teacher addStudent1(Student student){
        this.studentList.add(student);
        return this;
    }

    public Teacher addStudent2(Student student){
        this.studentList2.add(student);
        return this;
    }

}
