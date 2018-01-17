package com.manjusaka.test;

import com.excelread.annotation.Column;
import org.apache.commons.lang3.builder.ToStringBuilder;

/**
 * description:
 * ===>
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 9:47
 */
public class Test1 {
    @Column
    private String c;
    @Column
    private String v;
    @Column
    private String f;
    @Column
    private String d;
    @Column
    private String g;
    @Column
    private String e;
    @Column
    private String y;

    public String getC() {
        return c;
    }

    public void setC(String c) {
        this.c = c;
    }

    public String getV() {
        return v;
    }

    public void setV(String v) {
        this.v = v;
    }

    public String getF() {
        return f;
    }

    public void setF(String f) {
        this.f = f;
    }

    public String getD() {
        return d;
    }

    public void setD(String d) {
        this.d = d;
    }

    public String getG() {
        return g;
    }

    public void setG(String g) {
        this.g = g;
    }

    public String getE() {
        return e;
    }

    public void setE(String e) {
        this.e = e;
    }

    public String getY() {
        return y;
    }

    public void setY(String y) {
        this.y = y;
    }

    @Override
    public String toString() {
        return new ToStringBuilder(this)
                .append("c", c)
                .append("v", v)
                .append("f", f)
                .append("d", d)
                .append("g", g)
                .append("e", e)
                .append("y", y)
                .toString();
    }
}
