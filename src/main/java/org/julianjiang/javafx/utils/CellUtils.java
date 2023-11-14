package org.julianjiang.javafx.utils;

import org.apache.poi.ss.usermodel.*;

import java.util.Date;

public class CellUtils {
    public static void setCellValue(Cell cell, Object value) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        CellStyle cellStyle = cell.getCellStyle();

        if (value instanceof Double || value instanceof Float) {
            cell.setCellValue((Double) value);
            cell.setCellType(CellType.NUMERIC);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
            cell.setCellType(CellType.STRING);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
            cell.setCellType(CellType.BOOLEAN);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
            cell.setCellType(CellType.NUMERIC);
            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));
            cell.setCellStyle(cellStyle);
        } else {
            cell.setCellValue(value.toString());
            cell.setCellType(CellType.STRING);
        }
    }
}
