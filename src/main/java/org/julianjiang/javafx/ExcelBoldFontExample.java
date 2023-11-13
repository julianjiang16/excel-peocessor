package org.julianjiang.javafx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelBoldFontExample {
    public static void main(String[] args) {
        String filePath = "E:\\data\\excel-test\\tetsss.xlsx";
        int rowIndex = 0; // 要设置字体加粗的行索引
        int cellIndex = 0; // 要设置字体加粗的列索引

        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            Row row = sheet.createRow(rowIndex);
            Cell cell = row.createCell(cellIndex);
            cell.setCellValue("Hello, World!");

            // 创建字体对象并设置加粗
            Font font = workbook.createFont();
            font.setBold(true);

            // 创建单元格样式对象并应用字体
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);

            // 将样式应用于单元格
            cell.setCellStyle(cellStyle);

            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
            fos.close();

            System.out.println("Excel字体设置成功！");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
