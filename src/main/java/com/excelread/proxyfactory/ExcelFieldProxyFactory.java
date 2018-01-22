package com.excelread.proxyfactory;

import com.excelread.annotation.Column;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

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
     * @param data  excel数据
     * @param clazz 数据对象
     * @return 对象数据
     */
    public static Object getProxy(List<String> data, Class<?> clazz) throws IntrospectionException, IllegalAccessException, InstantiationException, InvocationTargetException {
        if (clazz == null) {
            return data;
        }
        Object obj = clazz.newInstance();
        int i = 0;
        for (Field field : clazz.getDeclaredFields()) {
            Column column = field.getAnnotation(Column.class);
            if (column != null) {
                int columnLen = column.length();
                String str = data.get(i);
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
