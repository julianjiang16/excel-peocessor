package org.julianjiang.javafx.processor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.julianjiang.javafx.utils.ExcelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelCopyRows {
    public static void main(String[] args) throws IOException {
        // 读取源Excel文件
        FileInputStream fis = new FileInputStream("E:\\data\\excel-test\\打单模板 - 副本.xlsx");
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // 创建目标Excel文件
        FileOutputStream fos = new FileOutputStream("E:\\data\\excel-test\\target.xlsx");
        Workbook targetWorkbook = new XSSFWorkbook();
        Sheet targetSheet = targetWorkbook.createSheet("Sheet1");

        // 拷贝前N行
        int n = 5;
        for (int i = 0; i < n; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Row targetRow = targetSheet.createRow(i);
                copyRow(row, targetRow);
            }
        }

        // 复制合并单元格
        int numMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++) {
            CellRangeAddress cra = sheet.getMergedRegion(i);
            targetSheet.addMergedRegion(cra);
            setMergeRegionWidth(sheet, cra, targetSheet);
        }

        // 保存目标Excel文件
        targetWorkbook.write(fos);
        fos.close();
        workbook.close();
    }

    private static void copyRow(Row srcRow, Row destRow) {
        short lastCellNum = srcRow.getLastCellNum();
        for (short cellNum = srcRow.getFirstCellNum(); cellNum < lastCellNum; cellNum++) {
            Cell srcCell = srcRow.getCell(cellNum);
            if (srcCell != null) {
                CellStyle style = srcCell.getCellStyle();
                final CellStyle copyCellStyle = ExcelUtils.getCopyCellStyle(style, srcRow.getSheet().getWorkbook(), destRow.getSheet().getWorkbook(), false);
                final Cell cell = destRow.createCell(cellNum);

                final float width = srcRow.getSheet().getColumnWidthInPixels(cellNum);
                destRow.getSheet().setColumnWidth(cellNum, Math.round(width / 256 * 256));
                cell.setCellStyle(copyCellStyle);
                if (srcCell.getCellType() == CellType.STRING) {
                    destRow.getCell(cellNum).setCellValue(srcCell.getStringCellValue());
                } else if (srcCell.getCellType() == CellType.NUMERIC) {
                    destRow.getCell(cellNum).setCellValue(srcCell.getNumericCellValue());
                } else if (srcCell.getCellType() == CellType.BOOLEAN) {
                    destRow.getCell(cellNum).setCellValue(srcCell.getBooleanCellValue());
                } else if (srcCell.getCellType() == CellType.FORMULA) {
                    destRow.getCell(cellNum).setCellFormula(srcCell.getCellFormula());
                }
            }
        }
    }

    private static void setMergeRegionWidth(Sheet srcSheet, CellRangeAddress cra, Sheet targetSheet) {
        int startCol = cra.getFirstColumn();
        int endCol = cra.getLastColumn();

        // 计算最宽的单元格的宽度
        float maxWidth = 0;
        for (int col = startCol; col <= endCol; col++) {
            float columnWidth = srcSheet.getColumnWidthInPixels(col);
            maxWidth = maxWidth + columnWidth;
        }

        // 设置合并的单元格的宽度
        targetSheet.setColumnWidth(startCol, Math.round(maxWidth));
    }
}
