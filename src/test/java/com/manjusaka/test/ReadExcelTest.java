package com.manjusaka.test;

import com.excelread.readexcel.Excel;

import java.io.File;
import java.io.IOException;
import java.util.Date;

/**
 * description:
 * ===>
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 10:10
 */
public class ReadExcelTest {

    public static void main(String[] arg) {
        for (int i = 0; i < 1; ++i) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    Date date = new Date();
                    String path = "d:/excel/1.xlsx";
                    File file = new File(path);
                    try {
                        System.out.println("读取全部" + Excel.readExcel(file, User.class));
                        System.out.println("读取指定sheet表" + Excel.readExcelByIndex(file, User.class, 1).size());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }).start();
        }
    }
}
