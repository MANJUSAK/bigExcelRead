package com.excel.writeexcel;

import com.excel.util.WriteExcelUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * description:
 * ===>导出excel入口类
 * company
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-06-13 14:11
 * @version V1.0.0
 */
public class WExcel {
    private final static String SUFFIX = "xlsx";

    /**
     * @param file      导出excel临时文件
     * @param list      写入到excel数据
     * @param titleList 写入excel导航栏数据
     * @param title     excel标题
     * @return File
     * @throws IOException io异常
     */
    public static File writeExcel(File file, List<List<String>> list, List<String> titleList, String title) throws IOException {
        String fileName = file.getName().toLowerCase();
        if (!fileName.endsWith(SUFFIX)) {
            throw new IOException("file format is not xlsx");
        }
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook();
            WriteExcelUtil.createExcel(wb, list, titleList, title).write(new BufferedOutputStream(fileOut));
        } finally {
            assert fileOut != null;
            fileOut.close();
        }
        return new File(file.getPath());
    }
}
