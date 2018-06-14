package com.excel.util;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

/**
 * description:
 * ===>写入数据到excel表工具类
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 17:33
 * @version V1.1.0
 */
public class WriteExcelUtil {
    private static final int LEN = 10000;

    /**
     * 创建excel表格
     * 将表单数据写入excel表格中
     *
     * @param wb        wb
     * @param list      写入到表格数据
     * @param titleList 标题栏数据
     * @param title     标题名称
     * @return XSSFWorkbook
     */
    public static XSSFWorkbook createExcel(XSSFWorkbook wb, List<List<String>> list, List<String> titleList, String title) {
        return createHeaderStyle(wb, list, titleList, title);
    }

    /**
     * 创建excel表头样式表格
     *
     * @param list      写入到表格数据
     * @param titleList 导航栏数据
     * @param title     标题名称
     * @return XSSFWorkbook
     */
    private static XSSFWorkbook createHeaderStyle(XSSFWorkbook wb, List<List<String>> list, List<String> titleList, String title) {
        XSSFCellStyle style = wb.createCellStyle();
        //水平居中
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        //设置单元格背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        //设置上下左右边框
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        // 自动换行
        style.setWrapText(true);
        XSSFFont font = wb.createFont();
        // 设置字体颜色
        font.setColor(HSSFColor.BLACK.index);
        // 设置字体样式
        font.setFontName("微软雅黑");
        // 设置字体大小
        font.setFontHeightInPoints((short) 11);
        // 设置字体粗细
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        // 把字体应用到当前的样式中
        style.setFont(font);
        return writeHeaderCellExcel(wb, style, list, titleList, title);
    }


    /**
     * 创建excel表头单元格
     * 将表单数据写入excel表格中
     *
     * @param wb        wb
     * @param style     style
     * @param list      导出数据
     * @param titleList 导航栏数据
     * @param title     标题名称
     * @return XSSFWorkbook
     */
    @SuppressWarnings("unchecked")
    private static XSSFWorkbook writeHeaderCellExcel(XSSFWorkbook wb, XSSFCellStyle style, List<List<String>> list, List<String> titleList, String title) {
        List data;
        int size = list.size();
        XSSFSheet sheet;
        CellRangeAddress cra;
        //判断数据是否超过一万条
        if (size > LEN) {
            //算出sheet表数
            int i = (size + LEN) / LEN;
            for (int k = 1; k <= i; ++k) {
                //判断是否为最后一张sheet表
                if (k < i) {
                    //截取写入到分表的数据
                    data = list.subList(((k - 1) * LEN), k * LEN);
                    //创建一张excel表
                    sheet = wb.createSheet("sheet" + k);
                    //合并单元格（第一行，第1-20列）
                    cra = new CellRangeAddress(0, 0, 0, titleList.size() - 1);
                    //在sheet里增加合并单元格
                    sheet.addMergedRegion(cra);
                    //获取第一行 写入标题名称
                    //设置列宽
                    sheet.setDefaultColumnWidth(35);
                    //创建表头
                    XSSFRow row = sheet.createRow(0);
                    //设置行高
                    row.setHeight((short) (16 * 25));
                    XSSFCell cell = row.createCell(0);
                    cell.setCellValue(title);
                    cell.setCellStyle(style);
                    writeHeaderCell(sheet, cell, style, titleList);
                    wb = createCellStyle(wb, sheet, data);
                } else {
                    data = list.subList(((k - 1) * LEN), size);
                    //创建一张excel表
                    sheet = wb.createSheet("sheet" + k);
                    cra = new CellRangeAddress(0, 0, 0, titleList.size() - 1);
                    //在sheet里增加合并单元格
                    sheet.addMergedRegion(cra);
                    //设置列宽
                    sheet.setDefaultColumnWidth(35);
                    //创建表头
                    XSSFRow row = sheet.createRow(0);
                    //设置行高
                    row.setHeight((short) (16 * 25));
                    XSSFCell cell = row.createCell(0);
                    cell.setCellValue(title);
                    cell.setCellStyle(style);
                    writeHeaderCell(sheet, cell, style, titleList);
                    wb = createCellStyle(wb, sheet, data);
                }
            }
            return wb;
        } else {
            //创建一张excel表
            sheet = wb.createSheet("sheet");
            cra = new CellRangeAddress(0, 0, 0, titleList.size() - 1);
            //在sheet里增加合并单元格
            sheet.addMergedRegion(cra);
            //设置列宽
            sheet.setDefaultColumnWidth(35);
            //创建表头
            XSSFRow row = sheet.createRow(0);
            //设置行高
            row.setHeight((short) (16 * 25));
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(title);
            cell.setCellStyle(style);
            //写入字段名称
            writeHeaderCell(sheet, cell, style, titleList);
            return createCellStyle(wb, sheet, list);
        }
    }

    /**
     * excel表头信息
     *
     * @param sheet     sheet
     * @param cell      cell
     * @param style     style
     * @param titleList 导航栏数据
     */
    private static void writeHeaderCell(XSSFSheet sheet, XSSFCell cell, XSSFCellStyle style, List<String> titleList) {
        XSSFRow row = sheet.createRow(1);
        // 创建表头单元格
        for (int i = 0, len = titleList.size(); i < len; ++i) {
            cell = row.createCell(i);
            cell.setCellValue(titleList.get(i));
            cell.setCellStyle(style);
        }
    }

    /**
     * 创建excel单元格样式
     * 将表单数据写入excel表格中
     *
     * @param wb    表属性
     * @param sheet 表
     * @param list  写入到表格数据
     * @return XSSFWorkbook
     */
    private static XSSFWorkbook createCellStyle(XSSFWorkbook wb, XSSFSheet sheet, List<List<String>> list) {
        XSSFCellStyle style = wb.createCellStyle();
        //水平居中
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        //设置上下左右边框
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        //设置自动换行
        //style.setWrapText(true);
        Font font = wb.createFont();
        // 设置字体颜色
        font.setColor(HSSFColor.BLACK.index);
        // 设置字体样式
        font.setFontName("微软雅黑");
        // 设置字体大小
        font.setFontHeightInPoints((short) 9);
        // 设置字体粗细
        font.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式中
        style.setFont(font);
        return nursery(wb, sheet, style, list);
    }

    /**
     * 创建excel单元格内容
     * 将表单数据写入excel表格中
     *
     * @param wb    表属性
     * @param sheet 表
     * @param style 表样式
     * @param list  写入到表格数据
     * @return XSSFWorkbook
     */
    private static XSSFWorkbook nursery(XSSFWorkbook wb, XSSFSheet sheet, XSSFCellStyle style, List<List<String>> list) {
        for (int i = 0, length = list.size(); i < length; ++i) {
            XSSFRow row = sheet.createRow(2 + i);
            //写入数据到sheet表中
            writeData(row, style, list.get(i));
        }
        return wb;
    }

    /**
     * 写入数据到excel表中
     *
     * @param row   row
     * @param style style
     * @param list  行数据
     */
    private static void writeData(XSSFRow row, XSSFCellStyle style, List<String> list) {
        XSSFCell cell;
        // 创建表内容单元格
        for (int i = 0, len = list.size(); i < len; ++i) {
            cell = row.createCell(i);
            cell.setCellValue(list.get(i));
            cell.setCellStyle(style);
        }
    }
}
