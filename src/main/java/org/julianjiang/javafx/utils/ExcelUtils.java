package org.julianjiang.javafx.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExcelUtils {

    public static void copyFile(File templateFile, int preRows, int lastRows, File outputFile, List<Map<String, Object>> data) throws IOException, InvalidFormatException {
        Workbook workbook = new XSSFWorkbook(templateFile);
        Sheet sheet = workbook.getSheetAt(0);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet();

        // 复制前N行
        for (int i = 0; i < preRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Row outputRow = outputSheet.createRow(i);
                copyRow(workbook, sheet, outputWorkbook, outputSheet, row, outputRow);
            }
        }

        // 写入数据
        int startRowIndex = outputSheet.getLastRowNum();
        for (int i = 0; i < data.size(); i++) {
            Row row = outputSheet.createRow(startRowIndex + i);
            Map<String, Object> rowData = data.get(i);
            int columnIndex = 0;
            for (String key : rowData.keySet()) {
                Cell cell = row.createCell(columnIndex);
                cell.setCellValue(rowData.get(key).toString());
                columnIndex++;
            }
        }


        // 复制后NN行
        int totalRows = outputSheet.getLastRowNum() + 1;
        for (int i = totalRows; i < lastRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Row outputRow = outputSheet.createRow(i);
                copyRow(workbook, sheet, outputWorkbook, outputSheet, row, outputRow);
            }
        }

        // 保存输出Excel文件
        FileOutputStream fileOut = new FileOutputStream(outputFile);
        outputWorkbook.write(fileOut);
        fileOut.close();

        workbook.close();
        outputWorkbook.close();
    }

    private static void copyRow(Workbook inputWorkbook, Sheet inputSheet, Workbook outputWorkbook, Sheet outputSheet, Row inputRow, Row outputRow) {
        for (Cell cell : inputRow) {
            Cell outputCell = outputRow.createCell(cell.getColumnIndex());
            outputCell.setCellValue(getCellValue(inputWorkbook, cell));

            CellStyle cellStyle = cell.getCellStyle();
            CellStyle outputCellStyle = outputWorkbook.createCellStyle();
            outputCellStyle.cloneStyleFrom(cellStyle);
            outputCell.setCellStyle(outputCellStyle);
        }

        // 处理合并单元格
        for (CellRangeAddress mergedRegion : inputSheet.getMergedRegions()) {
            if (mergedRegion.isInRange(inputRow.getRowNum(), inputRow.getRowNum())) {
                CellRangeAddress outputMergedRegion = new CellRangeAddress(
                        outputRow.getRowNum(), outputRow.getRowNum(),
                        mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                outputSheet.addMergedRegion(outputMergedRegion);
            }
        }
    }

    private static String getCellValue(Workbook workbook, Cell cell) {
        String cellValue = "";
        if (cell.getCellType() == CellType.STRING) {
            cellValue = cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            cellValue = String.valueOf(cell.getNumericCellValue());
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == CellType.FORMULA) {
            cellValue = String.valueOf(cell.getCellFormula());
        } else if (cell.getCellType() == CellType.BLANK) {
            cellValue = "";
        } else if (cell.getCellType() == CellType.ERROR) {
            cellValue = String.valueOf(cell.getErrorCellValue());
        }
        return cellValue;
    }
}