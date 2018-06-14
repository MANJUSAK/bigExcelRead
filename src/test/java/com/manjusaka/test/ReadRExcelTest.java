package com.manjusaka.test;

import com.excel.readexcel.RExcel;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * description:
 * ===>
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-16 10:10
 */
public class ReadRExcelTest {
    public static void main(String[] arg) {
        Date date = new Date();
        String path = "D:\\test\\test.xlsx";
        File file = new File(path);
        try {
            List<?> data = RExcel.readExcel(file, User.class);
            System.out.println("读取全部===>" + data);
            System.out.println("读取数量===>" + data.size());
            //System.out.println("读取指定sheet表===>" + RExcel.readExcelByIndex(file, User.class, 1).size());
        } catch (IOException e) {
            e.printStackTrace();
        }
        long t = System.currentTimeMillis();
        System.out.println("运行时间为：" + (t - date.getTime()) + "ms");
    }

    @Test
    public void test() {
        System.out.println(getMsg());

    }

    private int getMsg() {
        String path = "C:\\Users\\Lenovo\\Desktop\\各地区每日导入银海情况.xlsx";
        File file = new File(path);
        try {
            System.out.println("读取全部===>" + RExcel.readExcel(file, User.class).size());
            System.out.println("读取指定sheet表===>" + RExcel.readExcelByIndex(file, User.class, 1).size());
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

  /*  public static void main(String[] args) throws Exception {
        ReadExcelUtil excelUtil = new ReadExcelUtil();
        String filename = "D:\\test\\11.xlsx";
        System.out.println("-- 程序开始 --");
        long time_1 = System.currentTimeMillis();
        excelUtil.processOneSheet(filename);
        long time_2 = System.currentTimeMillis();
        System.out.println(ReadExcelUtil.dataList);
        System.out.println(ReadExcelUtil.dataList.size());
        String[] data = ReadExcelUtil.dataList.get(0).toString().split(",");
        for (int i = 0; i < data.length; ++i) {
            System.out.println(data[i]);
        }
        System.out.println("-- 程序结束1 --");
        System.out.println("-- 耗时1 --" + (time_2 - time_1) / 1000 + "s");
    }*/
}
