package org.julianjiang.javafx.processor;

import javafx.scene.control.Alert;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.julianjiang.javafx.Context;
import org.julianjiang.javafx.component.AlertComponent;
import org.julianjiang.javafx.utils.ExcelUtils;

import java.io.*;
import java.nio.file.Paths;
import java.util.*;

public class ExcelProcessor {

    public static Pair<String, Boolean> validate(Context context) {
        if (CollectionUtils.isEmpty(context.getData())) {
            return Pair.create("未选择明细数据文件 !!!", false);
        }
        if (CollectionUtils.isEmpty(context.getAllocation())) {
            return Pair.create("未选择分单条件 !!!", false);
        }
        if (Objects.isNull(context.getTemplateFile())) {
            return Pair.create("未选择输出模板文件 !!!", false);
        }
        if (StringUtils.isBlank(context.getOutputPath())) {
            return Pair.create("未选择输出文件夹 !!!", false);
        }

        return Pair.create("", true);
    }


    public static void outputExcel(Context context) throws IOException, InvalidFormatException {

        // 输出excel
        Pair<List<Cell>, List<Cell>> preLastCell = extractRowsFromExcel(context.getTemplateFile(), context.getPreRows(), context.getLastRows());
        // 分单数据
        Map<String, List<Map<String, Object>>> groupData = groupDataByAllocation(context.getData(), context.getAllocation());
        Alert alert = AlertComponent.buildAlert("处理中...", "有" + groupData.keySet().size() + "个文件需要生成，请耐心等待...");
        alert.show();
        for (Map.Entry<String, List<Map<String, Object>>> entity : groupData.entrySet()) {
//            writeExcel(entity.getKey(), entity.getValue(), preLastCell, context);
            String fileExtension = ".xlsx"; // 文件后缀
            String filePath = Paths.get(context.getOutputPath(), entity.getKey() + fileExtension).toString();
            ExcelUtils.copyFile(context.getTemplateFile(), context.getPreRows(), context.getLastRows(), new File(filePath), entity.getValue());
            // 先生成1个文件
            break;
        }
        if (alert.isShowing()) {
            alert.close();
        }
    }


    public static void writeExcel(String fileName, List<Map<String, Object>> data, Pair<List<Cell>, List<Cell>> preLastCell, Context context) throws IOException {
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Sheet1");
        // 写入前N行cell

        Workbook workbook = new XSSFWorkbook(new FileInputStream(context.getTemplateFile()));
        Sheet sheet = workbook.getSheetAt(0);
        List<Cell> preRowCells = preLastCell.getKey();
        for (int i = 0; i < preRowCells.size(); i++) {
            Cell cell = preRowCells.get(i);
            Row row = outputSheet.getRow(cell.getRowIndex());
            if (row == null) {
                row = outputSheet.createRow(cell.getRowIndex());
            }
            Cell newCell = row.createCell(cell.getColumnIndex());
            setCellValue(cell, newCell, outputWorkbook);

            // 处理合并单元格
            mergeCell(sheet, cell, outputSheet);
        }
        // 写入数据
        int startRowIndex = preRowCells.isEmpty() ? 0 : preRowCells.get(preRowCells.size() - 1).getRowIndex() + 1;
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
        // 写入后N行cell
        List<Cell> lastRowCells = preLastCell.getValue();
        int lastRowIndex = startRowIndex + data.size();
        for (int i = 0; i < lastRowCells.size(); i++) {
            Cell cell = lastRowCells.get(i);
            Row row = outputSheet.getRow(lastRowIndex + i);
            if (row == null) {
                row = outputSheet.createRow(lastRowIndex + i);
            }
            Cell newCell = row.createCell(cell.getColumnIndex());
            setCellValue(cell, newCell, outputWorkbook);
            // 处理合并单元格
            mergeCell(sheet, cell, outputSheet);
        }
        // 调整列宽
        for (int i = 0; i < outputSheet.getRow(0).getLastCellNum(); i++) {
            outputSheet.autoSizeColumn(i);
        }
        // 保存Excel文件
        String fileExtension = ".xlsx"; // 文件后缀
        String filePath = Paths.get(context.getOutputPath(), fileName + fileExtension).toString();
        FileOutputStream fileOut = new FileOutputStream(filePath);
        outputWorkbook.write(fileOut);
        fileOut.close();
        outputWorkbook.close();
    }

    public static void mergeCell(Sheet sheet, Cell cell, Sheet outputSheet) {
        // 处理合并单元格
        for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
            if (mergedRegion.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                boolean isAlreadyMerged = false;
                for (CellRangeAddress outputMergedRegion : outputSheet.getMergedRegions()) {
                    if (outputMergedRegion.getFirstRow() == mergedRegion.getFirstRow()
                            && outputMergedRegion.getFirstColumn() == mergedRegion.getFirstColumn()
                            && outputMergedRegion.getLastRow() == mergedRegion.getLastRow()
                            && outputMergedRegion.getLastColumn() == mergedRegion.getLastColumn()) {
                        isAlreadyMerged = true;
                        break;
                    }
                }
                if (!isAlreadyMerged) {
                    outputSheet.addMergedRegion(mergedRegion);
                }
            }
        }
    }


    public static void setCellValue(Cell old, Cell now, Workbook outputWorkbook) {
        if (old != null && now != null) {
            CellType cellType = old.getCellType();
            switch (cellType) {
                case STRING:
                    now.setCellValue(old.getStringCellValue());
                    break;
                case NUMERIC:
                    now.setCellValue(old.getNumericCellValue());
                    break;
                case BOOLEAN:
                    now.setCellValue(old.getBooleanCellValue());
                    break;
                case BLANK:
                    now.setCellValue("");
                    break;
            }
        }

        // 复制单元格样式
        CellStyle cellStyle = old.getCellStyle();
        CellStyle outputCellStyle = outputWorkbook.createCellStyle();
        outputCellStyle.cloneStyleFrom(cellStyle);
        now.setCellStyle(outputCellStyle);
    }


    public static Object getCellValue(Cell currentCell) {
        Object cellValue = "";
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
        return cellValue;
    }


    public static Map<String, List<Map<String, Object>>> groupDataByAllocation(List<Map<String, Object>> data, List<String> allocation) {
        // 创建一个Map，用于存储分类结果
        Map<String, List<Map<String, Object>>> groupedData = new HashMap<>();
        // 遍历数据列表
        for (Map<String, Object> item : data) {
            // 根据分类字段的值获取对应的键
            StringBuilder allocationKey = new StringBuilder();
            for (String allocationField : allocation) {
                allocationKey.append(item.get(allocationField)).append("-");
            }

            // 根据分类键将数据放入对应的列表中
            List<Map<String, Object>> group = groupedData.getOrDefault(allocationKey.toString(), new ArrayList<>());
            group.add(item);
            groupedData.put(allocationKey.toString(), group);
        }
        // 将分类结果转换为列表形式
        return groupedData;
    }


    public static Pair<List<Cell>, List<Cell>> extractRowsFromExcel(File templateFile, int preRows, int lastRows) throws IOException {
        List<Cell> preRowCells = new ArrayList<>();
        List<Cell> lastRowCells = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(templateFile));
        Sheet sheet = workbook.getSheetAt(0);
        int totalRows = sheet.getLastRowNum() + 1;
        // 抽取前N行单元格
        for (int i = 0; i < Math.min(preRows, totalRows); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    preRowCells.add(cell);
                }
            }
        }
        // 抽取最后X行单元格
        for (int i = Math.max(totalRows - lastRows, 0); i < totalRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    lastRowCells.add(cell);
                }
            }
        }
        // 保持单元格顺序与原Excel一致
        preRowCells.sort((c1, c2) -> {
            int rowComparison = Integer.compare(c1.getRowIndex(), c2.getRowIndex());
            if (rowComparison == 0) {
                return Integer.compare(c1.getColumnIndex(), c2.getColumnIndex());
            }
            return rowComparison;
        });
        lastRowCells.sort((c1, c2) -> {
            int rowComparison = Integer.compare(c1.getRowIndex(), c2.getRowIndex());
            if (rowComparison == 0) {
                return Integer.compare(c1.getColumnIndex(), c2.getColumnIndex());
            }
            return rowComparison;
        });

        return Pair.create(preRowCells, lastRowCells);
    }
}
