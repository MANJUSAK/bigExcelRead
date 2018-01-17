package com.excelread.service.lmpl;

import com.excelread.exception.ReadExcelException;
import com.excelread.proxyfactory.ExcelFieldProxyFactory;
import com.excelread.service.GetExcelDataService;
import com.excelread.util.AbstractReadExcelUtil;

import java.beans.IntrospectionException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

/**
 * description:
 * ===>获取excel数据接口实现类
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 17:36
 * @version V1.1.0
 */
public class GetExcelDataServiceImpl implements GetExcelDataService {

    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream excel文件
     * @param clazz       数据对象
     * @return list
     * @throws RuntimeException 异常
     */
    @SuppressWarnings("Duplicates")
    @Override
    public List<?> getExcelDataService(InputStream inputStream, final Class<?> clazz) throws ReadExcelException {
        final List<Object> list = new ArrayList<>();
        AbstractReadExcelUtil rxus = new AbstractReadExcelUtil() {
            @Override
            public void getRows(List<String> rowList) {
                try {
                    list.add(ExcelFieldProxyFactory.getProxy(rowList, clazz));
                } catch (IntrospectionException | InvocationTargetException | InstantiationException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        };
        try {
            rxus.process(inputStream);
        } catch (Exception e) {
            throw new ReadExcelException(e.getMessage());
        }
        return list;
    }

    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream excel文件
     * @param clazz       数据对象
     * @param sheetIndex  sheet表索引
     * @return list
     * @throws ReadExcelException 异常
     */
    @SuppressWarnings("Duplicates")
    @Override
    public List<?> getExcelDataByIndexService(InputStream inputStream, final Class<?> clazz, int sheetIndex) throws ReadExcelException {
        final List<Object> list = new ArrayList<>();
        AbstractReadExcelUtil rxu = new AbstractReadExcelUtil() {
            @Override
            public void getRows(List<String> rowList) {
                try {
                    list.add(ExcelFieldProxyFactory.getProxy(rowList, clazz));
                } catch (IntrospectionException | InvocationTargetException | InstantiationException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        };
        try {
            rxu.process(inputStream, sheetIndex);
        } catch (Exception e) {
            throw new ReadExcelException(e.getMessage());
        }
        return list;
    }
}
