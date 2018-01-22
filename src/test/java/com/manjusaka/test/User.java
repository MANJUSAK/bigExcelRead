package com.manjusaka.test;

import com.excelread.annotation.Column;
import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;

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
    @Column(length = 2)
    private String school;

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

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;

        if (!(o instanceof User)) return false;

        User user = (User) o;

        return new EqualsBuilder()
                .append(id, user.id)
                .append(name, user.name)
                .append(gender, user.gender)
                .append(age, user.age)
                .append(groupName, user.groupName)
                .append(clazz, user.clazz)
                .append(school, user.school)
                .isEquals();
    }

    @Override
    public int hashCode() {
        return new HashCodeBuilder(17, 37)
                .append(id)
                .append(name)
                .append(gender)
                .append(age)
                .append(groupName)
                .append(clazz)
                .append(school)
                .toHashCode();
    }

    @Override
    public String toString() {
        final StringBuffer sb = new StringBuffer("User{");
        sb.append("id='").append(id).append('\'');
        sb.append(", name='").append(name).append('\'');
        sb.append(", gender='").append(gender).append('\'');
        sb.append(", age='").append(age).append('\'');
        sb.append(", groupName='").append(groupName).append('\'');
        sb.append(", clazz='").append(clazz).append('\'');
        sb.append(", school='").append(school).append('\'');
        sb.append('}');
        return sb.toString();
    }
}
