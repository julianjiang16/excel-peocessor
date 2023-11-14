package org.julianjiang.javafx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import static org.julianjiang.javafx.utils.ExcelUtils.copyMergedRegions;
import static org.julianjiang.javafx.utils.ExcelUtils.extraPic;
import static org.julianjiang.javafx.utils.ExcelUtils.getCopyCellStyle;
import static org.julianjiang.javafx.utils.ExcelUtils.setCellValue;

public class Main {
    public static void main(String[] args) throws IOException {
        String inputFilePath = "E:\\data\\excel-test\\template.xlsx";
        String outputFilePath = "E:\\data\\excel-test\\template1.xlsx";

        try (Workbook workbookInput = WorkbookFactory.create(new File(inputFilePath));
             Workbook workbookOutput = new XSSFWorkbook();
             FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {

            Sheet sheetInput = workbookInput.getSheetAt(0); // 获取第一个工作表
            Sheet sheetOutput = workbookOutput.createSheet("Output"); // 创建新的工作表
            Iterator<Row> rowIterator = sheetInput.iterator();

            copyMergedRegions(sheetInput, sheetOutput);
            while (rowIterator.hasNext()) {
                Row rowInput = rowIterator.next();
                Row rowOutput = sheetOutput.createRow(sheetOutput.getLastRowNum() + 1);
                rowOutput.setRowStyle(rowInput.getRowStyle());
                for (int i = 0; i < rowInput.getLastCellNum(); i++) { // 遍历每一列
                    Cell cellInput = rowInput.getCell(i);
                    Cell cellOutput = rowOutput.createCell(i);
                    if (cellInput != null) {
                        setCellValue(cellInput, cellOutput);
                        cellOutput.setCellStyle(getCopyCellStyle(cellInput.getCellStyle(), workbookInput, workbookOutput));
                    }
                }
            }

//            addPic(sheetOutput, workbookOutput, imgPath);
            extraPic(sheetInput, sheetOutput, workbookOutput);
            workbookOutput.write(fileOut);
        }
    }

}


