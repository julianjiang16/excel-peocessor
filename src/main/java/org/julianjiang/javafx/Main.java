package org.julianjiang.javafx;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

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
            for (int columnIndex = 0; columnIndex < sheetOutput.getRow(0).getLastCellNum(); columnIndex++) {
                sheetOutput.autoSizeColumn(columnIndex);
            }
            String imgPath = "E:\\data\\excel-test\\test.jpg";
            addPic(sheetOutput, workbookOutput, imgPath);
            workbookOutput.write(fileOut);
        }
    }


    private static void addPic(Sheet sheet, Workbook workbook, String imageFilePath) throws IOException {
        byte[] imageBytes = IOUtils.toByteArray(new FileInputStream(imageFilePath));
        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_JPEG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(0); // 图片起始列
        anchor.setRow1(0); // 图片起始行

        // 插入图片
        Picture picture = drawing.createPicture(anchor, pictureIdx);
        picture.resize(); // 自适应图片大小
    }


    private static void setCellValue(Cell srcCell, Cell destCell) {
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

    private static CellStyle getCopyCellStyle(CellStyle sourceStyle, Workbook sourceWorkbook, Workbook destinationWorkbook) {
        CellStyle destinationStyle = destinationWorkbook.createCellStyle();
        destinationStyle.setAlignment(sourceStyle.getAlignment());
        destinationStyle.setVerticalAlignment(sourceStyle.getVerticalAlignment());
        destinationStyle.setDataFormat(sourceStyle.getDataFormat());
        Font sourceFont = sourceWorkbook.getFontAt(sourceStyle.getFontIndex());
        Font destinationFont = destinationWorkbook.createFont();
        copyFont(sourceFont, destinationFont);
        destinationStyle.setFont(destinationFont);
        destinationStyle.setFillForegroundColor(sourceStyle.getFillForegroundColorColor());
        destinationStyle.setFillPattern(sourceStyle.getFillPattern());
        destinationStyle.setBorderTop(sourceStyle.getBorderTop());
        destinationStyle.setBorderBottom(sourceStyle.getBorderBottom());
        destinationStyle.setBorderLeft(sourceStyle.getBorderLeft());
        destinationStyle.setBorderRight(sourceStyle.getBorderRight());
        destinationStyle.setWrapText(sourceStyle.getWrapText());
        destinationStyle.setRotation(sourceStyle.getRotation());
        destinationStyle.setIndention(sourceStyle.getIndention());
        destinationStyle.setLocked(sourceStyle.getLocked());
        destinationStyle.setHidden(sourceStyle.getHidden());
        destinationStyle.cloneStyleFrom(sourceStyle);
        return destinationStyle;
    }

    // 复制字体的方法
    private static void copyFont(Font sourceFont, Font destinationFont) {
        destinationFont.setFontHeight(sourceFont.getFontHeight());
        destinationFont.setFontName(sourceFont.getFontName());
        destinationFont.setItalic(sourceFont.getItalic());
        destinationFont.setBold(sourceFont.getBold());
        destinationFont.setColor(sourceFont.getColor());
        destinationFont.setStrikeout(sourceFont.getStrikeout());
        destinationFont.setTypeOffset(sourceFont.getTypeOffset());
        destinationFont.setUnderline(sourceFont.getUnderline());
        destinationFont.setCharSet(sourceFont.getCharSet());
    }

    private static void copyMergedRegions(Sheet srcSheet, Sheet destSheet) {
        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = srcSheet.getMergedRegion(i);
            destSheet.addMergedRegion(mergedRegion);
        }
    }
}


