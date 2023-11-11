package org.julianjiang.javafx.processor;

import com.google.common.collect.Lists;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class FileProcessor {

    public static Pair<ArrayList<String>, List<Map<String, Object>>> readExcel(InputStream inputStream) throws IOException {
        List<Map<String, Object>> data = new ArrayList<>();

        ArrayList<String> headers = Lists.newArrayList();
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row headerRow = sheet.getRow(0);
        int columnCount = headerRow.getLastCellNum();

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row currentRow = sheet.getRow(rowIndex);
            Map<String, Object> rowData = new HashMap<>();

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                Cell currentCell = currentRow.getCell(columnIndex);
                Object cellValue = "";
                String header = headerRow.getCell(columnIndex).getStringCellValue();
                if (currentCell != null) {
                    CellType cellType = currentCell.getCellType();
                    switch (cellType) {
                        case STRING:
                            cellValue = currentCell.getStringCellValue();
                            break;
                        case NUMERIC:
                            cellValue = currentCell.getNumericCellValue();
                            break;
                        case BOOLEAN:
                            cellValue = currentCell.getBooleanCellValue();
                            break;
                        case BLANK:
                            cellValue = "";
                            break;
                    }
                }

                rowData.put(header, cellValue);
            }
            data.add(rowData);
        }
        headers.addAll(data.get(0).keySet());
        workbook.close();
        return Pair.create(headers, data);
    }
}
