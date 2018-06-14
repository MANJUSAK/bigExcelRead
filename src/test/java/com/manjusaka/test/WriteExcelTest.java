package com.manjusaka.test;

import com.excel.readexcel.RExcel;
import com.excel.writeexcel.WExcel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * description:
 * ===>
 * company
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-06-13 18:14
 * @version V1.0.0
 */
public class WriteExcelTest {
    public static void main(String[] args) {
        try {
            boolean bl = GetOsNameUtil.getOsName();
            StringBuilder sb = new StringBuilder();
            if (bl) {
                sb.append("/home/");
            } else {
                sb.append("d:");
            }
            String path = "/test";
            sb.append(path);
            File mkdirs = new File(sb.toString());
            if (!mkdirs.exists()) {
                mkdirs.mkdirs();
            }
            sb.append("/test.xlsx");
            File file = new File(sb.toString());
            if (file.exists()) {
                file.delete();
                file.createNewFile();
            } else {
                file.createNewFile();
            }
            List<String> titleList = new ArrayList<>();
            titleList.add("列名1");
            titleList.add("列名2");
             /* titleList.add("列名3");
            titleList.add("列名4");
            titleList.add("列名5");
            titleList.add("列名6");
            titleList.add("列名7");
            titleList.add("列名8");*/
            List<String> list1 = new ArrayList<>();
            list1.add("值1");
            list1.add("值2");
            /*list1.add("值3");
            list1.add("值4");
            list1.add("值5");
            list1.add("值6");
            list1.add("值7");
            list1.add("值8");*/
            List<List<String>> list = new ArrayList<>();
            for (int i = 0; i < 100; ++i) {
                list.add(list1);
            }
            String title = "标题名称";
            File files = WExcel.writeExcel(file, list, titleList, title);
            System.out.println(files.getPath());
            List<?> data = RExcel.readExcel(files, User.class);
            System.out.println("读取全部===>" + data);
            System.out.println("读取全部===>" + data.size());
            //System.out.println("读取指定sheet表===>" + RExcel.readExcelByIndex(file, User.class, 1).size());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
