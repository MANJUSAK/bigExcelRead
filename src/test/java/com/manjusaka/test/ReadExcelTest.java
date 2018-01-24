package com.manjusaka.test;

import com.excelread.readexcel.Excel;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Random;

/**
 * description:
 * ===>
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 10:10
 */
public class ReadExcelTest {
    public static void main(String[] arg) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                Date date = new Date();
                String path = "d:/excel/8.xlsx";
                File file = new File(path);
                try {
                    System.out.println("读取全部===>" + Excel.readExcel(file, User.class));
                    //System.out.println("读取指定sheet表===>" + Excel.readExcelByIndex(file, User.class, 1).size());
                } catch (IOException e) {
                    e.printStackTrace();
                }
                long t = System.currentTimeMillis();
                System.out.println("运行时间为：" + (t - date.getTime()) + "ms");
            }
        }).start();
        new Thread(new Runnable() {
            @Override
            public void run() {
                Date date = new Date();
                String path = "d:/excel/1.xlsx";
                File file = new File(path);
                try {
                    System.out.println("读取全部===>" + Excel.readExcel(file, User.class));
                    //System.out.println("读取指定sheet表===>" + Excel.readExcelByIndex(file, User.class, 1).size());
                } catch (IOException e) {
                    e.printStackTrace();
                }
                long t = System.currentTimeMillis();
                System.out.println("运行时间为：" + (t - date.getTime()) + "ms");
            }
        }).start();
        new Thread(new Runnable() {
            @Override
            public void run() {
                Date date = new Date();
                String path = "d:/excel/2.xlsx";
                File file = new File(path);
                try {
                    System.out.println("读取全部===>" + Excel.readExcel(file, User.class));
                    //System.out.println("读取指定sheet表===>" + Excel.readExcelByIndex(file, User.class, 1).size());
                } catch (IOException e) {
                    e.printStackTrace();
                }
                long t = System.currentTimeMillis();
                System.out.println("运行时间为：" + (t - date.getTime()) + "ms");
            }
        }).start();
       /* for (int i = 0; i < 10; ++i) {
            for (int j = 1; j <= i; ++j) {
                System.out.print(i + "*" + j + "=" + i * j + "\t");
            }
            System.out.println();
        }*/
    }

    @Test
    public void test() {
        System.out.println(getMsg());

    }

    private int getMsg() {
        String path = "d:/excel/8.xlsx";
        File file = new File(path);
        try {
            System.out.println("读取全部===>" + Excel.readExcel(file, User.class).size());
            System.out.println("读取指定sheet表===>" + Excel.readExcelByIndex(file, User.class, 1).size());
        } catch (IOException e) {
            e.printStackTrace();
        }
        int code = new Random().nextInt(1000000);
        if (code < 100000) {
            System.out.println(code);
            return getMsg();
        } else {
            return code;
        }
    }
}
