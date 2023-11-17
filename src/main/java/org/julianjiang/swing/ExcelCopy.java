package org.julianjiang.swing;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.julianjiang.javafx.utils.ExcelUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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
        copyMergedRegions(srcSheet, destSheet);
        for (int i = srcSheet.getFirstRowNum(); i <= ExcelUtils.getLastRowWithData(srcSheet); i++) {
            Row srcRow = srcSheet.getRow(i);
            Row destRow = destSheet.createRow(i);
            copyRow(srcRow, destRow, srcWorkbook, destWorkbook);
        }
    }

    private static void copyMergedRegions(Sheet srcSheet, Sheet destSheet) {
        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = srcSheet.getMergedRegion(i);
            destSheet.addMergedRegion(mergedRegion);
        }
    }

    private static void copyCellStyle(Workbook srcWorkbook, CellStyle sourceStyle, CellStyle targetStyle) {
        targetStyle.setAlignment(sourceStyle.getAlignment());
        targetStyle.setVerticalAlignment(sourceStyle.getVerticalAlignment());
        targetStyle.setDataFormat(sourceStyle.getDataFormat());

        targetStyle.setFont(cloneFont(srcWorkbook, sourceStyle.getFontIndex()));

        targetStyle.setFillForegroundColor(sourceStyle.getFillForegroundColor());
        targetStyle.setFillBackgroundColor(sourceStyle.getFillBackgroundColor());
        targetStyle.setFillPattern(sourceStyle.getFillPattern());
        targetStyle.setBorderTop(sourceStyle.getBorderTop());
        targetStyle.setBorderBottom(sourceStyle.getBorderBottom());
        targetStyle.setBorderLeft(sourceStyle.getBorderLeft());
        targetStyle.setBorderRight(sourceStyle.getBorderRight());
        targetStyle.setWrapText(sourceStyle.getWrapText());
        targetStyle.setRotation(sourceStyle.getRotation());
        targetStyle.setIndention(sourceStyle.getIndention());
        targetStyle.setLocked(sourceStyle.getLocked());
        targetStyle.setHidden(sourceStyle.getHidden());
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
        newFont.setBold(sourceFont.getBold());
        return newFont;
    }

    private static void copyRow(Row srcRow, Row destRow, Workbook srcWorkbook, Workbook destWorkbook) {
        destRow.setHeight(srcRow.getHeight());
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            Cell srcCell = srcRow.getCell(j);
            Cell destCell = destRow.createCell(j);
            copyCell(srcCell, destCell, srcWorkbook, destWorkbook);
        }
    }

    private static void copyCell(Cell srcCell, Cell destCell, Workbook srcWorkbook, Workbook destWorkbook) {
        if (srcCell != null) {
            CellStyle srcCellStyle = srcCell.getCellStyle();
            if (srcCellStyle != null) {
                CellStyle destCellStyle = destWorkbook.createCellStyle();
                copyCellStyle(srcWorkbook, srcCellStyle, destCellStyle);
                destCell.setCellStyle(destCellStyle);
            }
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
                    destCell.setBlank();
                    break;
                case FORMULA:
                    destCell.setCellFormula(srcCell.getCellFormula());
                    break;
                default:
                    break;
            }
        }
    }
}
