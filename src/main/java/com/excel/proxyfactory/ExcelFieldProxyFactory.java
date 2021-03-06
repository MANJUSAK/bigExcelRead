package com.excel.proxyfactory;

import com.excel.annotation.Column;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * description:
 * ===>读取到的excel数据映射到实体对象代理工厂类
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 14:26
 * @version V1.1.0
 */
public class ExcelFieldProxyFactory {
    /**
     * 字段数据映射
     *
     * @param data  <code>excel数据</code>
     * @param clazz <code>数据对象</code>
     * @return <code>对象数据</code>
     */
    public static Object getProxy(String[] data, Class<?> clazz) throws IntrospectionException, IllegalAccessException, InstantiationException, InvocationTargetException {
        if (clazz == null) {
            return data;
        }
        Object obj = clazz.newInstance();
        int i = 0;
        int len = data.length;
        for (Field field : clazz.getDeclaredFields()) {
            if (i >= len) {
                break;
            }
            Column column = field.getAnnotation(Column.class);
            if (column != null) {
                int columnLen = column.length();
                String str = data[i];
                ++i;
                PropertyDescriptor pd = new PropertyDescriptor(field.getName(), clazz);
                Method method = pd.getWriteMethod();
                if (column.nonempty()) {
                    if (str != null) {
                        method.invoke(obj, str);
                    }
                } else if (columnLen != 0) {
                    if (str.length() <= columnLen) {
                        method.invoke(obj, str);
                    }
                } else {
                    method.invoke(obj, str);
                }
            }
        }
        return obj;
    }
}
