package org.julianjiang;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.julianjiang.javafx.utils.ExcelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelCopy {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("E:\\data\\excel-test\\template.xlsx");

        FileOutputStream out = new FileOutputStream("E:\\data\\excel-test\\template1.xlsx");
        Workbook srcWorkbook = new XSSFWorkbook(fis);
        Sheet srcSheet = srcWorkbook.getSheetAt(0);

        Workbook destWorkbook = new XSSFWorkbook();
        Sheet destSheet = destWorkbook.createSheet("Output");

        copySheet(srcSheet, destSheet, srcWorkbook, destWorkbook);

        destWorkbook.write(out);
        out.close();
        srcWorkbook.close();
        destWorkbook.close();
        fis.close();

    }

    private static void copySheet(Sheet srcSheet, Sheet destSheet, Workbook srcWorkbook, Workbook destWorkbook) {
        for (int i = ExcelUtils.getLastRowWithData(srcSheet); i >= 0; i--) {
            Row srcRow = srcSheet.getRow(i);
            Row destRow = destSheet.createRow(i);
            copyRow(srcRow, destRow, srcWorkbook, destWorkbook);
        }
        copyMergedRegions(srcSheet, destSheet);
    }

    private static void copyMergedRegions(Sheet srcSheet, Sheet destSheet) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        for (int i = srcSheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = srcSheet.getMergedRegion(i);
            mergedRegions.add(new CellRangeAddress(mergedRegion.getFirstRow(), mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn()));
        }
        for (CellRangeAddress mergedRegion : mergedRegions) {
            destSheet.addMergedRegion(mergedRegion);
        }
    }


    private static void copyCellStyle(Workbook srcWorkbook, CellStyle sourceStyle, CellStyle targetStyle) {
        targetStyle.setDataFormat(srcWorkbook.createDataFormat().getFormat(sourceStyle.getDataFormatString()));

        targetStyle.setFillBackgroundColor(sourceStyle.getFillBackgroundColor());
        targetStyle.setFillForegroundColor(sourceStyle.getFillForegroundColor());
        targetStyle.setFont(cloneFont(srcWorkbook, sourceStyle.getFontIndex()));
        targetStyle.setRotation((short) sourceStyle.getRotation());
    }

    private static Font cloneFont(Workbook srcWorkbook, int sourceFontIndex) {
        Font sourceFont = srcWorkbook.getFontAt(sourceFontIndex);

        Font newFont = srcWorkbook.createFont();
        newFont.setFontName(sourceFont.getFontName());
        newFont.setFontHeight(sourceFont.getFontHeight());
        newFont.setColor(sourceFont.getColor());
        newFont.setItalic(sourceFont.getItalic());
        newFont.setStrikeout(sourceFont.getStrikeout());
        newFont.setTypeOffset(sourceFont.getTypeOffset());
        newFont.setUnderline(sourceFont.getUnderline());
        newFont.setFontHeightInPoints(sourceFont.getFontHeightInPoints());
        // 设置加粗属性
        newFont.setBold(sourceFont.getBold());

        return newFont;
    }

    private static boolean isCellOnDiagonal(Cell cell) {
        Row row = cell.getRow();
        int rowIndex = row.getRowNum();
        int columnIndex = cell.getColumnIndex();

        // 判断行索引和列索引是否相等
        if (rowIndex == columnIndex) {
            return true;
        }

        // 判断行索引是否大于列索引
        if (rowIndex > columnIndex) {
            return true;
        }

        return false;
    }

    private static void setCellDiagonalLine(Cell cell) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();

        // 设置斜线样式
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        cellStyle.setBorderDiagonal(BorderStyle.THIN);
//        cellStyle.setDiagonalBorderColor(IndexedColors.BLACK.getIndex());
//
        cell.setCellStyle(cellStyle);
    }

    private static void copyRow(Row srcRow, Row destRow, Workbook srcWorkbook, Workbook destWorkbook) {
        int lastCellNum = srcRow.getLastCellNum();
        destRow.setHeight(srcRow.getHeight());
        for (int j = 0; j < lastCellNum; j++) {
            Cell srcCell = srcRow.getCell(j);
            if (srcCell == null)
                continue;

            // 获取当前样式信息，并确保样式属于目标工作簿
            CellStyle style = destWorkbook.createCellStyle();
            copyCellStyle(srcWorkbook, srcCell.getCellStyle(), style);

            // 复制单元格及其样式
            Cell destCell = destRow.createCell(j);
            if (isCellOnDiagonal(srcCell)) {
                setCellDiagonalLine(destCell);
            }
            // 设置单元格
            int cellWidth = srcCell.getSheet().getColumnWidth(srcCell.getColumnIndex());
            destCell.getSheet().setColumnWidth(destCell.getColumnIndex(), cellWidth);
            destCell.setCellStyle(style);
            switch (srcCell.getCellType()) {
                case STRING:
                    destCell.setCellValue(srcCell.getStringCellValue());
                    break;
                case NUMERIC:
                    destCell.setCellValue(srcCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    destCell.setCellValue(srcCell.getBooleanCellValue());
                    break;
                case BLANK:
                    destCell.setCellType(CellType.BLANK);
                    break;
                case FORMULA:
                    destCell.setCellFormula(srcCell.getCellFormula());
                    break;
                default:
                    System.out.println("Unknown cell type: " + srcCell.getCellType());
            }
        }
    }
}
