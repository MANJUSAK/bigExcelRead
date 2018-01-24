package com.excelread.readexcel;

import com.excelread.exception.ReadExcelException;
import com.excelread.service.GetExcelDataService;
import com.excelread.service.lmpl.GetExcelDataServiceImpl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
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
     * @param file  excel文件
     * @param clazz 数据对象：1.数据对象实体字段必须按照excel列进行命名，否则会出现字段与数据不一致（如：excel第一列对应实体从上往下第一字段属性）
     *              2.实体字段属性必须加上此注解@Column，否则会出现字段映射不到数据的情况
     * @return excel数据
     * @throws ReadExcelException 组件异常
     * @throws IOException        io异常
     */
    public static List<?> readExcel(File file, Class<?> clazz) throws ReadExcelException, IOException {
        FileInputStream fis = null;
        List<?> list;
        try {
            fis = new FileInputStream(file);
            GetExcelDataService getExcelDataService = new GetExcelDataServiceImpl();
            list = getExcelDataService.getExcelDataService(fis, clazz);
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
     * @param file  excel文件
     * @param clazz 数据对象：1.数据对象实体字段必须按照excel列进行命名，否则会出现字段与数据不一致（如：excel第一列对应实体从上往下第一字段属性）
     *              2.实体字段属性必须加上此注解@Column，否则会出现字段映射不到数据的情况
     * @return excel数据
     * @throws ReadExcelException 组件异常
     * @throws IOException        io异常
     */
    public static List<?> readExcelByIndex(File file, Class<?> clazz, int sheetIndex) throws ReadExcelException, IOException {
        FileInputStream fis = null;
        List<?> list;
        try {
            fis = new FileInputStream(file);
            GetExcelDataService getExcelDataService = new GetExcelDataServiceImpl();
            list = getExcelDataService.getExcelDataByIndexService(fis, clazz, sheetIndex);
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
