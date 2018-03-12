package com.excelread.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

/**
 * description:
 * ===>解析excel数据
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 17:33
 * @version V1.1.0
 */
public abstract class AbstractReadExcelUtil extends DefaultHandler {
    /**
     * 共享字符串表
     */
    private SharedStringsTable sst;
    /**
     * 上一次的内容
     */
    private String lastContents;
    private boolean nextIsString;
    /**
     * 读取每行数据
     */
    private List<String> rowList = new ArrayList<>();
    /**
     * 当前列
     */
    private int curCol = 0;
    /**
     * 日期标志
     */
    private boolean dateFlag;
    /**
     * 数字标志
     */
    private boolean numberFlag;
    /**
     * 空值标志
     */
    private boolean isTElement;

    /**
     * 遍历工作簿中所有的电子表格
     *
     * @param inputStream <code>文件</code>
     * @throws Exception <code>异常</code>
     */
    public void process(InputStream inputStream) throws Exception {
        OPCPackage pkg = OPCPackage.open(inputStream);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        Iterator<InputStream> sheets = r.getSheetsData();
        while (sheets.hasNext()) {
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
        pkg.close();
    }

    /**
     * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
     *
     * @param inputStream <code>文件</code>
     * @param sheetId     <code>sheet表索引</code>
     * @throws Exception <code>异常</code>
     */
    public void process(InputStream inputStream, int sheetId) throws Exception {
        OPCPackage pkg = OPCPackage.open(inputStream);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        // 根据 rId# 或 rSheet# 查找sheet
        InputStream sheet = r.getSheet("rId" + sheetId);
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();
        pkg.close();
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst = sst;
        parser.setContentHandler(this);
        return parser;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        // c => 单元格
        if ("c".equals(name)) {
            // 如果下一个元素是 SST 的索引，则将nextIsString标记为true
            String cellType = attributes.getValue("t");
            nextIsString = "s".equals(cellType);
            // 日期格式
            String cellDateType = attributes.getValue("s");
            dateFlag = "1".equals(cellDateType);
            String cellNumberType = attributes.getValue("s");
            numberFlag = "2".equals(cellNumberType);
        }
        // 当元素为t时,不管是不是空值都设置为true
        if ("t".equals(name)) {
            isTElement = true;
        } else {
            isTElement = true;
        }
        // 置空
        lastContents = "";
    }

    @Override
    public void endElement(String uri, String localName, String name)
            throws SAXException {
        // 根据SST的索引值的到单元格的真正要存储的字符串
        // 这时characters()方法可能会被调用多次
        if (nextIsString) {
            try {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            } catch (Exception ignored) {
            }
        }
        // t元素也包含字符串
        if (isTElement) {
            String value = lastContents.trim();
            rowList.add(curCol, value);
            curCol++;
            isTElement = false;
            // v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
            // 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
        } else if ("v".equals(name)) {
            String value = lastContents.trim();
            value = "".equals(value) ? " " : value;
            try {
                // 日期格式处理
                if (dateFlag) {
                    Date date = HSSFDateUtil.getJavaDate(Double.valueOf(value));
                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                    value = dateFormat.format(date);
                }
                // 数字类型处理
                if (numberFlag) {
                    BigDecimal bd = new BigDecimal(value);
                    value = bd.setScale(3, BigDecimal.ROUND_UP).toString();
                }
            } catch (Exception ignored) {
            }
            rowList.add(curCol, value);
            curCol++;
        } else {
            // 如果标签名称为 row ，这说明已到行尾，调用 getRows() 方法
            if ("row".equals(name)) {
                try {
                    getRows(rowList);
                } catch (Exception ignored) {
                }
                rowList.clear();
                curCol = 0;
            }
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        // 得到单元格内容的值
        lastContents += new String(ch, start, length);
    }

    /**
     * 获取行数据回调
     *
     * @param rowList <code>行数据</code>
     */
    public abstract void getRows(List<String> rowList);
}
