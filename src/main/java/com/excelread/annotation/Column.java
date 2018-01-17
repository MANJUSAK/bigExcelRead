package com.excelread.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * description:
 * ===>excel映射字段注解
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 14:14
 * @version V1.1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface Column {
    /**
     * 字段内容长度
     *
     * @return 长度
     */
    int length() default 0;

    /**
     * 字段是否为空
     *
     * @return true/false
     */
    boolean nonempty() default false;

}
