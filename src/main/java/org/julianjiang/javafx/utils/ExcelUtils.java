package org.julianjiang.javafx.utils;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Triple;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.julianjiang.javafx.model.SheetContext;
import org.julianjiang.javafx.processor.ExcelProcessor;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

public class ExcelUtils {

    // 指定范围合并样式应用
    public static void copyMergedCellStyles(Sheet srcSheet, Sheet destSheet, int srcRowIndex, int destRowIndex, int rows) {
        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = srcSheet.getMergedRegion(i);
            int mergedRegionFirstRow = mergedRegion.getFirstRow();
            int mergedRegionLastRow = mergedRegion.getLastRow();

            if (mergedRegionFirstRow >= srcRowIndex && mergedRegionFirstRow < srcRowIndex + rows) {
                int destRegionFirstRow = destRowIndex + (mergedRegionFirstRow - srcRowIndex);
                int destRegionLastRow = destRegionFirstRow + (mergedRegionLastRow - mergedRegionFirstRow);

                CellRangeAddress destRegion = new CellRangeAddress(destRegionFirstRow, destRegionLastRow,
                        mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                destSheet.addMergedRegion(destRegion);
            }
        }
    }


    public static void extraPicV2(Sheet inputSheet, Sheet outSheet, Workbook outWorkbook) {
        Drawing<?> drawing = inputSheet.getDrawingPatriarch();
        for (Object obj : drawing) {
            if (obj instanceof Picture) {
                Picture picture = (Picture) obj;
                ClientAnchor anchor = picture.getClientAnchor();
                int row = anchor.getDx1();
                int col = anchor.getDy1();
                addPic(outSheet, outWorkbook, picture.getPictureData().getData(), row, col, anchor);
            }
        }
    }

    public static void addPic(Sheet sheet, Workbook workbook, byte[] imageBytes, int row, int col, ClientAnchor anchor) {
        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor clientAnchor = helper.createClientAnchor();
        clientAnchor.setRow1(anchor.getRow1());
        clientAnchor.setCol1(anchor.getCol1());
        clientAnchor.setAnchorType(anchor.getAnchorType());
        clientAnchor.setDx1(anchor.getDx1());
        clientAnchor.setDy1(anchor.getDy1());
        clientAnchor.setDx2(anchor.getDx2());
        clientAnchor.setDy2(anchor.getDy2());
        clientAnchor.setRow2(anchor.getRow2());

        clientAnchor.setCol2(anchor.getCol2());

        // 插入图片
        Picture picture = drawing.createPicture(clientAnchor, pictureIdx);
        picture.resize(1.0); // 自适应图片大小
    }

    public static void extraPic(Sheet inputSheet, Sheet outSheet, Workbook outWorkbook) {
        Drawing<?> drawing = inputSheet.getDrawingPatriarch();
        if (drawing == null) {
            return;
        }
        // 遍历所有的图形对象
        Set<Picture> copiedPictures = new HashSet<>();
        final HashMap<Picture, Triple<Integer, Integer, Picture>> keyRowCol = Maps.newHashMap();
        for (Object obj : drawing) {
            if (obj instanceof Picture) {
                Picture picture = (Picture) obj;
                ClientAnchor anchor = picture.getClientAnchor();
                int row = anchor.getRow1();
                int col = anchor.getCol1();
                if (copiedPictures.contains(picture)) {
                    // 元素相同则判断 起始位置
                    final Triple<Integer, Integer, Picture> triple = keyRowCol.get(picture);
                    if (triple.getLeft().compareTo(row) > 0) {
                        keyRowCol.put(picture, Triple.of(row, col, picture));
                    }
                    continue;
                }
                copiedPictures.add(picture);
                keyRowCol.put(picture, Triple.of(row, col, picture));
            }
        }

        keyRowCol.values().forEach(item -> {
            addPic(outSheet, outWorkbook, item.getRight().getPictureData().getData(), item.getLeft(), item.getMiddle(), null);
        });

    }


    public static void setCellValue(Cell srcCell, Cell destCell, SheetContext sheetContext) {
        switch (srcCell.getCellType()) {
            case STRING:
                final String matchedString = PatternUtils.getMatchedString(srcCell.getStringCellValue());
                // 带后缀
                if (StringUtils.isNotBlank(matchedString)) {
                    final Object val = sheetContext.getReplaceMap().get(matchedString);
                    if (Objects.isNull(val)) {
                        destCell.setCellValue("");
                        break;
                    }

                    if (val instanceof Double) {
                        final String temp = "#" + matchedString + "#";
                        if (srcCell.getStringCellValue().equals(temp)) {
                            destCell.setCellValue((double) val);
                            break;
                        }
                    }
                    destCell.setCellValue(srcCell.getStringCellValue().replace("#" + matchedString + "#", String.valueOf(val)));
                } else {
                    destCell.setCellValue(srcCell.getStringCellValue());
                }
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

    public static Map<CellStyle, CellStyle> cacheCellStyle = Maps.newHashMap();

    public static CellStyle getCopyCellStyle(CellStyle sourceStyle, Workbook sourceWorkbook, Workbook destinationWorkbook, boolean headerFlag) {

        if (cacheCellStyle.get(sourceStyle) != null) {
            return cacheCellStyle.get(sourceStyle);
        }
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
        destinationStyle.setRotation(sourceStyle.getRotation());
        destinationStyle.setIndention(sourceStyle.getIndention());
        destinationStyle.setLocked(sourceStyle.getLocked());
        destinationStyle.setHidden(sourceStyle.getHidden());
        destinationStyle.cloneStyleFrom(sourceStyle);
        destinationStyle.setWrapText(sourceStyle.getWrapText());
        cacheCellStyle.put(sourceStyle, destinationStyle);
        return destinationStyle;
    }

    // 复制字体的方法
    public static void copyFont(Font sourceFont, Font destinationFont) {
        destinationFont.setFontHeight(sourceFont.getFontHeight());
        destinationFont.setFontName(sourceFont.getFontName());
        destinationFont.setItalic(sourceFont.getItalic());
        destinationFont.setBold(sourceFont.getBold());
        destinationFont.setColor(sourceFont.getColor());
        destinationFont.setStrikeout(sourceFont.getStrikeout());
        destinationFont.setTypeOffset(sourceFont.getTypeOffset());
        destinationFont.setUnderline(sourceFont.getUnderline());
        destinationFont.setCharSet(sourceFont.getCharSet());
        destinationFont.setFontHeightInPoints(sourceFont.getFontHeightInPoints());
    }

    public static void copyMergedRegions(Sheet srcSheet, Sheet destSheet) {
        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = srcSheet.getMergedRegion(i);
            destSheet.addMergedRegion(mergedRegion);
        }
    }

    public static List<String> getReplaceNames(InputStream inputStream) throws IOException {
        List<String> data = new ArrayList<>();

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        for (int rowIndex = 0; rowIndex <= ExcelUtils.getLastRowWithData(sheet); rowIndex++) {
            Row currentRow = sheet.getRow(rowIndex);
            for (int columnIndex = 0; columnIndex < currentRow.getLastCellNum(); columnIndex++) {
                Cell currentCell = currentRow.getCell(columnIndex);
                try {
                    String cellValue = (String) ExcelProcessor.getCellValue(currentCell);
                    final String matchedString = PatternUtils.getMatchedString(cellValue);
                    if (StringUtils.isNotBlank(matchedString)) {
                        data.add(matchedString);
                    }
                } catch (Exception e) {
                    continue;
                }
            }
        }
        workbook.close();
        return data;
    }

    public static Pair<ArrayList<String>, List<Map<String, Object>>> readExcel(InputStream inputStream) throws IOException {
        List<Map<String, Object>> data = new ArrayList<>();

        ArrayList<String> headers = Lists.newArrayList();
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row headerRow = sheet.getRow(0);
        int columnCount = headerRow.getLastCellNum();

        for (int rowIndex = 1; rowIndex <= ExcelUtils.getLastRowWithData(sheet); rowIndex++) {
            Row currentRow = sheet.getRow(rowIndex);
            Map<String, Object> rowData = new HashMap<>();

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                Cell currentCell = currentRow.getCell(columnIndex);
                String header = headerRow.getCell(columnIndex).getStringCellValue();
                Object cellValue = ExcelProcessor.getCellValue(currentCell);

                rowData.put(header, cellValue);
            }
            data.add(rowData);
        }
        headers.addAll(data.get(0).keySet());
        workbook.close();
        return Pair.create(headers, data);
    }


    public static List<Object> extraHeader(Sheet sheet, int row) {

        ArrayList<Object> headers = Lists.newArrayList();

        Row headerRow = sheet.getRow(row - 1);
        int columnCount = headerRow.getLastCellNum();
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            Cell currentCell = headerRow.getCell(columnIndex);
            headers.add(getCellValue(currentCell));
        }
        return headers;
    }


    public static Object getCellValue(Cell srcCell) {
        if (srcCell == null) {
            return "";
        }
        switch (srcCell.getCellType()) {
            case STRING:
                return srcCell.getStringCellValue();
            case NUMERIC:
                return srcCell.getNumericCellValue();
            case BOOLEAN:
                return srcCell.getBooleanCellValue();
            case FORMULA:
                return srcCell.getCellFormula();
            case BLANK:
            default:
                return "";
        }
    }

    public static int getLastRowWithData(Sheet sheet) {
        int lastRow = sheet.getLastRowNum();
        int lastRowWithData = -1;
        for (int i = lastRow; i >= 0; i--) {
            boolean breakFlag = false;
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = row.getLastCellNum() - 1; j >= 0; j--) {
                    Cell cell = row.getCell(j);
                    if (cell != null && cell.getCellType() != CellType.BLANK) {
                        lastRowWithData = i;
                        breakFlag = true;
                        break;
                    }
                }
            }
            if (breakFlag) {
                break;
            }
        }
        return lastRowWithData;
    }
}
