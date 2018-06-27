package com.manjusaka.test;

import cn.afterturn.easypoi.excel.annotation.Excel;
import org.apache.commons.lang3.builder.ToStringBuilder;

import javax.validation.constraints.Max;
import javax.validation.constraints.Min;
import javax.validation.constraints.Pattern;

/**
 * description:
 * ===>
 * company
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-06-14 14:22
 * @version V1.0.0
 */
public class Company {

    @Excel(name = "id")
    private String id;
    @Excel(name = "name")
    @Pattern(regexp = "[\\u4E00-\\u9FA5]{2,5}", message = "姓名中文2-5位")
    private String name;
    @Excel(name = "age")
    @Min(0)
    @Max(100)
    private String age;
    @Excel(name = "gender")
    @Pattern(regexp = "^1\\d{10}$", message = "电话")
    private String gender;


    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }


    @Override
    public String toString() {
        return new ToStringBuilder(this)
                .append("id", id)
                .append("name", name)
                .append("gender", gender)
                .append("age", age)
                .toString();
    }
}
