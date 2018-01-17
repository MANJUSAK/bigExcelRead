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
public interface GetExcelDataService {
    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream excel文件
     * @param clazz       数据对象
     * @return list
     * @throws ReadExcelException 异常
     */
    List<?> getExcelDataService(InputStream inputStream, Class<?> clazz) throws ReadExcelException;

    /**
     * 获取excel数据业务处理方法
     *
     * @param inputStream excel文件
     * @param clazz       数据对象
     * @param sheetIndex  sheet表索引
     * @return list
     * @throws ReadExcelException 异常
     */
    List<?> getExcelDataByIndexService(InputStream inputStream, Class<?> clazz, int sheetIndex) throws ReadExcelException;
}
