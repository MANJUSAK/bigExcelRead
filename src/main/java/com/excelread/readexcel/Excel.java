package com.excelread.readexcel;

import com.excelread.exception.ReadExcelException;
import com.excelread.service.ExcelDataService;
import com.excelread.service.lmpl.ExcelDataServiceImpl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

/**
 * description:
 * ===>读取excel入口类
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 10:00
 * @version V1.1.0
 */
public class Excel {
    /**
     * 读取大数据excel
     *
     * @param file  <code>excel文件</code>
     * @param clazz <code>数据对象：</code>
     *              <code>1.数据对象实体字段必须按照excel列进行命名，否则会出现字段与数据不一致（如：excel第一列对应实体从上往下第一字段属性）</code>
     *              <code> 2.实体字段属性必须加上此注解@Column，否则会出现字段映射不到数据的情况</code>
     * @return <code> List<?> </code>
     * @throws ReadExcelException <code>组件异常</code>
     * @throws IOException        <code>io异常</code>
     */
    public static List<?> readExcel(File file, Class<?> clazz) throws ReadExcelException, IOException {
        FileInputStream fis = null;
        LinkedList<?> list;
        try {
            fis = new FileInputStream(file);
            ExcelDataService excelDataService = new ExcelDataServiceImpl();
            list = (LinkedList<?>) excelDataService.getExcelDataService(fis, clazz);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new ReadExcelException(e.getMessage());
        } finally {
            assert fis != null;
            fis.close();
        }
        return list;
    }

    /**
     * 读取大数据excel,指定sheet索引
     *
     * @param file  <code>excel文件</code>
     * @param clazz <code>数据对象：</code>
     *              <code>1.数据对象实体字段必须按照excel列进行命名，否则会出现字段与数据不一致（如：excel第一列对应实体从上往下第一字段属性）</code>
     *              <code> 2.实体字段属性必须加上此注解@Column，否则会出现字段映射不到数据的情况</code>
     * @return <code> List<?> </code>
     * @throws ReadExcelException <code>组件异常</code>
     * @throws IOException        <code>io异常</code>
     */
    public static List<?> readExcelByIndex(File file, Class<?> clazz, int sheetIndex) throws ReadExcelException, IOException {
        FileInputStream fis = null;
        LinkedList<?> list;
        try {
            fis = new FileInputStream(file);
            ExcelDataService excelDataService = new ExcelDataServiceImpl();
            list = (LinkedList<?>) excelDataService.getExcelDataByIndexService(fis, clazz, sheetIndex);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new ReadExcelException(e.getMessage());
        } finally {
            assert fis != null;
            fis.close();
        }
        return list;
    }
}
