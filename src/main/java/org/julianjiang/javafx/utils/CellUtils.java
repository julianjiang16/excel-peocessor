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
//            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("@"));
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
//            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));
            cell.setCellStyle(cellStyle);
        } else {
            cell.setCellValue(value.toString());
        }
    }
}
