package com.manjusaka.test;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;
import org.junit.Test;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * description:
 * ===>
 * company
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-06-14 13:34
 * @version V1.0.0
 */
public class WexcelTest {
    @Test
    public void test() throws IOException {
        // 获取导出excel指定模版
        TemplateExportParams params = new TemplateExportParams("d:/test/mb1.xlsx");
        // 设置sheetName，若不设置该参数，则使用得原本得sheet名称
        params.setSheetName("test");
        Map<String, Object> map = new HashMap<>();
        List<Map<String, Object>> mapList = new ArrayList<>();
        Map<String, Object> key = new HashMap<>();
        key.put("id", "id");
        key.put("name", "name");
        key.put("age", "age");
        key.put("gender", "gender");
        mapList.add(key);
        for (int i = 0; i < 20; ++i) {
            Map<String, Object> data = new HashMap<>();
            data.put("id", i);
            data.put("name", "姓名");
            data.put("age", 1 + i);
            data.put("gender", "1850851391" + i);
            mapList.add(data);
        }
        map.put("mapList", mapList);
        File file = new File("d:/test/test.xlsx");
        FileOutputStream out = new FileOutputStream(file);
        // 导出excel
        ExcelExportUtil.exportExcel(params, map).write(new BufferedOutputStream(out));
        ImportParams importParams = new ImportParams();
        importParams.setTitleRows(2);
        importParams.setNeedVerfiy(true);
        IExcelDataHandler<Company> handler = new CompanyHandler();
        handler.setNeedHandlerFields(new String[]{"id", "name", "age", "gender"});
        importParams.setDataHandler(handler);
        // 导入excel
        /*FileInputStream in = new FileInputStream(file);
        BufferedInputStream bin = new BufferedInputStream(in);*/
        ExcelImportResult<Company> data = ExcelImportUtil.importExcelMore(file, Company.class, importParams);
        System.out.println(data.getList());
        System.out.println(data.getList().size());
    }
}
