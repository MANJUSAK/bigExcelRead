package com.excelread.service;

import com.excelread.exception.ReadExcelException;

import java.io.InputStream;
import java.util.List;

/**
 * description:
 * ===>获取excel数据
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 17:33
 * @version V1.1.0
 */
public interface ExcelDataService {
    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream <code>excel文件</code>
     * @param clazz       <code>数据对象</code>
     * @return <code>List<?></code>
     * @throws ReadExcelException <code>异常</code>
     */
    List<?> getExcelDataService(InputStream inputStream, Class<?> clazz) throws ReadExcelException;

    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream <code>excel文件</code>
     * @param clazz       <code>数据对象</code>
     * @param sheetIndex  <code>sheet表索引</code>
     * @return <code>List<?></code>
     * @throws ReadExcelException <code>异常</code>
     */
    List<?> getExcelDataByIndexService(InputStream inputStream, Class<?> clazz, int sheetIndex) throws ReadExcelException;
}
