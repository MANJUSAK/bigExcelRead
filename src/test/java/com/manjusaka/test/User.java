package com.manjusaka.test;

import com.excel.annotation.Column;
import org.apache.commons.lang3.builder.ToStringBuilder;

/**
 * description:
 * ===>
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 11:25
 */
public class User implements java.io.Serializable {
    private static final long serialVersionUID = 1606964350747145937L;
    @Column
    private String id;
    @Column
    private String name;
    @Column
    private String gender;
    @Column
    private String age;
    @Column
    private String groupName;
    @Column
    private String clazz;
    @Column
    private String school;
    @Column
    private String value;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
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

    public String getGroupName() {
        return groupName;
    }

    public void setGroupName(String groupName) {
        this.groupName = groupName;
    }

    public String getClazz() {
        return clazz;
    }

    public void setClazz(String clazz) {
        this.clazz = clazz;
    }

    public String getSchool() {
        return school;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return new ToStringBuilder(this)
                .append("id", id)
                .append("name", name)
                .append("gender", gender)
                .append("age", age)
                .append("groupName", groupName)
                .append("clazz", clazz)
                .append("school", school)
                .append("value", value)
                .toString();
    }
}
