package com.lp.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class XlsUtil {
    /**
     * @param inputStream 这个inputStream文件可以来源于本地文件的流， 也可以来源与上传上来的文件的流，也就是MultipartFile的流，使用getInputStream()方法进行获取。
     * @param fileName    文件名
     */
    public void readExcelFile(InputStream inputStream, String fileName) {
        Workbook workbook = null;
        try {
            //判断什么类型文件
            if (fileName.endsWith(".xls")) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (fileName.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        if (workbook == null) return;

        //获取所有的工作表的的数量
        int sheetNum = workbook.getNumberOfSheets();
        System.out.println(sheetNum + "--->sheetNum");
        for (int i = 0; i < sheetNum; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet == null) continue;
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum == 0) continue;
            Row row;
            for (int j = 1; j <= lastRowNum; j++) {
                row = sheet.getRow(j);
                if (row == null) continue;

                short lastCellNum = row.getLastCellNum();
                for (int k = 0; k <= lastCellNum; k++) {
                    if (row.getCell(k) == null) continue;

                    row.getCell(k).setCellType(Cell.CELL_TYPE_STRING);
                    String res = row.getCell(k).getStringCellValue().trim();
                    System.out.println(res);
                }

            }
        }
    }

    /**
     *
     * @param outputStream 这个outputstream可以来自与文件的输出流也可以直接输出到response的getOutputStream()里面,然后用户就可以直接解析到你生产的excel文件了
     */
    public void writeExcel(OutputStream outputStream) {
        Workbook wb = new SXSSFWorkbook(100);
        Sheet sheet = wb.createSheet("sheet");
        Row row = sheet.createRow(0);
        for (int i = 0; i < 10; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("这是第" + i + "个cell");
        }
        try {
            wb.write(outputStream);
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
