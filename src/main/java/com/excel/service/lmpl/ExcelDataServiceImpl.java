package com.excel.service.lmpl;

import com.excel.exception.ReadExcelException;
import com.excel.proxyfactory.ExcelFieldProxyFactory;
import com.excel.service.ExcelDataService;
import com.excel.util.ReadExcelUtil;

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
        ReadExcelUtil excelUtil = new ReadExcelUtil();
        try {
            excelUtil.processOneSheet(inputStream);
        } catch (Exception e) {
            e.printStackTrace();
            throw new ReadExcelException(e.getMessage());
        }
        List data = ReadExcelUtil.dataList;
        for (Object aData : data) {
            try {
                list.add(ExcelFieldProxyFactory.getProxy(aData.toString().split(","), clazz));
            } catch (IntrospectionException | InvocationTargetException | InstantiationException | IllegalAccessException e) {
                throw new ReadExcelException(e.getMessage());
            }
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
        ReadExcelUtil excelUtil = new ReadExcelUtil();
        try {
            excelUtil.processOneSheet(inputStream, sheetIndex);
        } catch (Exception e) {
            e.printStackTrace();
            throw new ReadExcelException(e.getMessage());
        }
        List data = ReadExcelUtil.dataList;
        for (Object aData : data) {
            try {
                list.add(ExcelFieldProxyFactory.getProxy(aData.toString().split(","), clazz));
            } catch (IntrospectionException | InvocationTargetException | InstantiationException | IllegalAccessException e) {
                throw new ReadExcelException(e.getMessage());
            }
        }
        return list;
    }
}
