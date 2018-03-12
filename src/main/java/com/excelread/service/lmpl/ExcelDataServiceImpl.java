package com.excelread.service.lmpl;

import com.excelread.exception.ReadExcelException;
import com.excelread.proxyfactory.ExcelFieldProxyFactory;
import com.excelread.service.ExcelDataService;
import com.excelread.util.AbstractReadExcelUtil;

import java.beans.IntrospectionException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.LinkedList;
import java.util.List;

/**
 * description:
 * ===>获取excel数据接口实现类
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 17:36
 * @version V1.1.0
 */
public class ExcelDataServiceImpl implements ExcelDataService {

    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream <code>excel文件</code>
     * @param clazz       <code>数据对象</code>
     * @return <code>List<?></code>
     * @throws ReadExcelException <code>异常</code>
     */
    @SuppressWarnings("Duplicates")
    @Override
    public LinkedList<?> getExcelDataService(InputStream inputStream, final Class<?> clazz) throws ReadExcelException {
        final LinkedList<Object> list = new LinkedList<>();
        AbstractReadExcelUtil rxus = new AbstractReadExcelUtil() {
            @Override
            public void getRows(List<String> rowList) {
                try {
                    list.add(ExcelFieldProxyFactory.getProxy(rowList, clazz));
                } catch (IntrospectionException | InvocationTargetException | InstantiationException | IllegalAccessException e) {
                    throw new ReadExcelException(e.getMessage());
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
     * @param inputStream <code>excel文件</code>
     * @param clazz       <code>数据对象</code>
     * @param sheetIndex  <code>sheet表索引</code>
     * @return <code>List<?></code>
     * @throws ReadExcelException <code>异常</code>
     */
    @SuppressWarnings("Duplicates")
    @Override
    public LinkedList<?> getExcelDataByIndexService(InputStream inputStream, final Class<?> clazz, int sheetIndex) throws ReadExcelException {
        final LinkedList<Object> list = new LinkedList<>();
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
