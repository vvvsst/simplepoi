package org.simplepoi.test.entity;


import org.simplepoi.excel.annotation.ExcelField;

public class People {
    @ExcelField(name = "姓名", width = 15)
    private String name;
    @ExcelField(name = "年龄", width = 15)
    private Integer age;

    // one people can have multiple students

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
