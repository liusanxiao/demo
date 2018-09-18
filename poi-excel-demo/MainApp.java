package com.test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainApp {
    private static String path = "D:\\excel\\统计账款.xls";
    private static double maxValue = 0;
    private static int defaultSheet = 0;
    public static void main(String[] args) {
    	String inPath = System.getProperty("path");
    	String sheetIndexStr = System.getProperty("sheet");
    	int sheetIndex = sheetIndexStr == null ? defaultSheet : Integer.valueOf(sheetIndexStr) - 1;
    	if(inPath != null && !"".equals(inPath)){
    		System.out.println("输入的路径为:" + inPath);
    		path = inPath;
    	}
    	String maxValueStr = System.getProperty("maxValue");
    	if(maxValueStr == null ) {
    		System.out.println("请指定临界值");
    		return;
    	}
    	try {
    		maxValue = Double.valueOf(maxValueStr);
    		System.out.println("临界值为:" + maxValue);
    	} catch (Exception e) {
    		System.out.println("指定maxValue的值为正确的数值");
    		return;
    	}
        InputStream in = null;
        Workbook excel = null;
        Sheet sheet = null;
        int totalRow = 0;
        File file = new File(path);
        try {
            in = new FileInputStream(file);
            excel = createworkbook(in);
            in.close();
            sheet = excel.getSheetAt(sheetIndex);
            totalRow = sheet.getLastRowNum();
            Map<String, Double> map = new HashMap<String, Double>();
            Row row = null;
            String name = null;
            Double money = null;
            CellStyle style = excel.createCellStyle();
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            for (int i = 1; i <= totalRow; i++) {
                row = sheet.getRow(i);
                name = row.getCell(2).getStringCellValue();
                if(name == null || "".equals(name)){
                	continue;
                }
                name = name.trim();
                money = (Double) row.getCell(3).getNumericCellValue();
                map.put(name, map.get(name) == null ? money : map.get(name) + money);
                if(map.get(name) > maxValue) {
                	System.out.println(name + " 当前总额:" + map.get(name) + ",此行增加的值:" + money);
                	row.getCell(2).setCellStyle(style);
                	row.createCell(6).setCellValue(map.get(name));
                }
            }
            OutputStream out = new FileOutputStream(file);
            excel.write(out);
            out.flush();
            out.close();
            excel.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }
    
    public static Workbook createworkbook(InputStream inp) throws IOException,InvalidFormatException {
        if (!inp.markSupported()) {
            inp = new PushbackInputStream(inp, 8);
        }
        if (POIFSFileSystem.hasPOIFSHeader(inp)) {
            return new HSSFWorkbook(inp);
        }
        if (POIXMLDocument.hasOOXMLHeader(inp)) {
            return new XSSFWorkbook(OPCPackage.open(inp));
        }
        throw new IllegalArgumentException("你的excel版本目前poi解析不了");
    }
}
