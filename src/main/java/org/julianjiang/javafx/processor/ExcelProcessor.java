package org.julianjiang.javafx.processor;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.julianjiang.javafx.model.Context;
import org.julianjiang.javafx.model.ExcelTemplate;
import org.julianjiang.javafx.model.ProcessType;
import org.julianjiang.javafx.model.SheetContext;
import org.julianjiang.javafx.utils.CellUtils;
import org.julianjiang.javafx.utils.ExcelUtils;
import org.julianjiang.javafx.utils.FormulaUtils;
import org.julianjiang.javafx.utils.SimpleDateThreadLocal;

import javax.script.ScriptException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.julianjiang.javafx.Constants.*;
import static org.julianjiang.javafx.utils.ExcelUtils.extraPic;

public class ExcelProcessor {

    public static Pair<String, Boolean> validate(Context context) {
        if (CollectionUtils.isEmpty(context.getData())) {
            return Pair.create("未选择明细数据文件 !!!", false);
        }
        if (CollectionUtils.isEmpty(context.getAllocation())) {
            return Pair.create("未选择分单条件 !!!", false);
        }
        if (Objects.isNull(context.getExcelTemplate().getTemplateFile())) {
            return Pair.create("未选择输出模板文件 !!!", false);
        }
        if (StringUtils.isBlank(context.getOutputPath())) {
            return Pair.create("未选择输出文件夹 !!!", false);
        }

        /** 列名行号校验 */
        if (context.getExcelTemplate().getTitleRow() < 1) {
            return Pair.create("列名所处行号必须大于0 !!!", false);
        }

        /** 如果选择分类，则分类汇总行所处行号必填！！ */
        if (context.isTypeFlag()) {
            if (context.getExcelTemplate().getTypeRow() < 1) {
                return Pair.create("分类汇总行所处行号必须大于0 !!!", false);
            }
        }

        /** 校验明细数据行号  */
        if (context.getExcelTemplate().getDetailRow() < 1) {
            return Pair.create("明细数据行行号必须大于0 !!!", false);
        }

        return Pair.create("", true);
    }

    public static void main(String[] args) throws IOException, ScriptException {
        final ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setPreRows(5);
        excelTemplate.setLastRows(4);

        excelTemplate.setTypeRow(7);
        excelTemplate.setDetailRow(6);
        excelTemplate.setTitleRow(5);
        final File templateFile = new File("C:\\Users\\Administrator\\Desktop\\模板.xlsx");
        excelTemplate.setTemplateFile(templateFile);
        final Context context = new Context(excelTemplate);
        final File detailFile = new File("C:\\Users\\Administrator\\Desktop\\明细.xlsx");
        Pair<ArrayList<String>, List<Map<String, Object>>> dataPair = ExcelUtils.readExcel(new FileInputStream(detailFile));
        context.setAllocation(Lists.newArrayList("发货时间"));
        context.setTypeFlag(true);
        context.setOutputPath("C:\\Users\\Administrator\\Desktop");
        context.setData(dataPair.getValue());
        context.setReplaceNames(ExcelUtils.getReplaceNames(new FileInputStream(templateFile)));
        outputExcel(context);
    }

    public static void outputExcel(Context context) throws IOException, ScriptException {
        Workbook workbookInput = WorkbookFactory.create(new File(context.getExcelTemplate().getTemplateFile().getAbsolutePath()));
        Sheet sheetInput = workbookInput.getSheetAt(0);
        // 先画页头 并替换变量
        Workbook outputWorkbook = new XSSFWorkbook();
        // 画列名
        final List<Object> titleNames = ExcelUtils.extraHeader(sheetInput, context.getExcelTemplate().getTitleRow());
        // 分单数据 每个sheet页数据
        final Map<String, List<Map<String, Object>>> allocationDataMap = groupDataByAllocation(context.getData(), context.getAllocation());
        Row rowInput = sheetInput.getRow(context.getExcelTemplate().getDetailRow() - 1);
        for (Map.Entry<String, List<Map<String, Object>>> allocationData : allocationDataMap.entrySet()) {

            final SheetContext sheetContext = new SheetContext();
            // 处理普通替换字段
            handleSpecificField(context.getReplaceNames(), sheetContext, allocationData.getValue(), ProcessType.SAMPLE);

            // 需要重新创建sheet页
            final String sheetName = allocationData.getKey();
            Sheet outputSheet = outputWorkbook.createSheet(sheetName);
            copyTemplateHeader(sheetInput, outputSheet, context, workbookInput, outputWorkbook, sheetContext);
            // 分类数据
            Map<String, List<Map<String, Object>>> typeDataMap = groupDataByAllocation(allocationData.getValue(), Lists.newArrayList(TYPE_COLUMN_NAME));
            if (!context.isTypeFlag()) {
                // 不是分类则构造新的sheet页数据
                typeDataMap = Maps.newHashMap();
                typeDataMap.put(allocationData.getKey(), allocationData.getValue());
            }

            for (Map.Entry<String, List<Map<String, Object>>> typeData : typeDataMap.entrySet()) {
                String typeName = typeData.getKey();
                // 取模板中明细行样式
                final List<Map<String, Object>> detailDatas = typeData.getValue();
                int seq = 1;
                for (Map<String, Object> detailData : detailDatas) {
                    Row rowOutput = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                    for (int i = 0; i < titleNames.size(); i++) {
                        final Object cellData = detailData.get(titleNames.get(i));
                        Cell cellInput = rowInput.getCell(i);
                        final Cell cellOutPut = rowOutput.createCell(i);
                        cellOutPut.setCellStyle(ExcelUtils.getCopyCellStyle(cellInput.getCellStyle(), workbookInput, outputWorkbook, false));
                        if (SERIAL_NUM_COLUMN_NAME.equals(titleNames.get(i))) {
                            CellUtils.setCellValue(cellOutPut, seq++);
                        } else {
                            if (Objects.isNull(cellData)) {
                                CellUtils.setCellValue(cellOutPut, "");
                            } else {
                                CellUtils.setCellValue(cellOutPut, cellData);
                            }
                        }
                        setCellWidth(sheetInput, outputSheet, i);
                    }
                }
                // 汇总行
                if (context.isTypeFlag()) {
                    // 汇总行处理汇总字段
                    handleSpecificField(context.getReplaceNames(), sheetContext, typeData.getValue(), ProcessType.TYPE);
                    copyTemplateType(sheetInput, outputSheet, context, workbookInput, outputWorkbook, typeName, sheetContext);
                }
            }
            // 单个 sheet页的后续工作
            // 处理总计
            handleSpecificField(context.getReplaceNames(), sheetContext, allocationData.getValue(), ProcessType.SUM);
            handleSpecificField(context.getReplaceNames(), sheetContext, allocationData.getValue(), ProcessType.FORMULA);
            copyTemplateBottom(sheetInput, outputSheet, context, workbookInput, outputWorkbook, sheetContext);
            // 处理单sheet 图片
            extraPic(sheetInput, outputSheet, outputWorkbook);

            // todo jcj 先1个 sheet
            setPageSize(sheetInput, outputSheet);
            outputSheet.setAutobreaks(true);
            break;
        }

        String fileExtension = ".xlsx"; // 文件后缀
        String filePath = Paths.get(context.getOutputPath(), "测试输出" + fileExtension).toString();
        // 保存输出Excel文件
        FileOutputStream fileOut = new FileOutputStream(filePath);
        outputWorkbook.write(fileOut);
        fileOut.close();

        workbookInput.close();
        outputWorkbook.close();

        // 最后处理图片附件

        // 输出excel
      /*  final Pair<List<Row>, List<Row>> preLastRows = extractRowsFromExcel(context.getExcelTemplate().getTemplateFile(), context.getExcelTemplate().getPreRows(), context.getExcelTemplate().getLastRows());
        // 分单数据
        Map<String, List<Map<String, Object>>> groupData = groupDataByAllocation(context.getData(), context.getAllocation());
        Alert alert = AlertComponent.buildAlert("处理中...", "有" + groupData.keySet().size() + "个文件需要生成，请耐心等待...");
        alert.show();
        for (Map.Entry<String, List<Map<String, Object>>> entity : groupData.entrySet()) {
//            writeExcel(entity.getKey(), entity.getValue(), preLastCell, context);
            String fileExtension = ".xlsx"; // 文件后缀
            String filePath = Paths.get(context.getOutputPath(), entity.getKey() + fileExtension).toString();
            ExcelUtils.copyFile(context.getExcelTemplate().getTemplateFile(), context.getExcelTemplate().getPreRows(), context.getExcelTemplate().getLastRows(), new File(filePath), entity.getValue());
            // 先生成1个文件
            break;
        }
        if (alert.isShowing()) {
            alert.close();
        }*/
    }

    public static void handleSpecificField(List<String> replaceNames, SheetContext sheetContext, List<Map<String, Object>> data, ProcessType type) throws ScriptException {
        switch (type) {
            case SAMPLE:
                for (String replaceName : replaceNames) {
                    if (!replaceName.endsWith(SUM_SUFFIX) && !replaceName.endsWith(TYPE_SUFFIX) && !replaceName.endsWith(FORMULA_SUFFIX)) { // 就是简单替换
                        final Object replaceVal = data.get(0).get(replaceName);
                        sheetContext.getReplaceMap().put(replaceName, replaceVal);
                    }
                }
                break;
            case SUM:
                for (String replaceName : replaceNames) {
                    if (replaceName.endsWith(SUM_SUFFIX)) { // 总计
                        // 转换为真实字段
                        final String realFieldName = replaceName.replaceAll(SUM_SUFFIX, "");
                        final double sum = data.stream().map(i -> i.get(realFieldName)).mapToDouble(i -> (double) i).sum();
                        sheetContext.getReplaceMap().put(replaceName, sum);
                    }
                }
                break;
            case TYPE:
                for (String replaceName : replaceNames) {
                    if (replaceName.endsWith(TYPE_SUFFIX)) { //汇总
                        final String realFieldName = replaceName.replaceAll(TYPE_SUFFIX, "");
                        // 先检测目标数据类型
                        final Object val = data.get(0).get(realFieldName);
                        if (val instanceof String) {
                            String replaceVal = (String) val;
                            sheetContext.getReplaceMap().put(replaceName, replaceVal);
                            continue;
                        }

                        if (val instanceof Double || val instanceof Float) {
                            final double sum = data.stream().map(i -> i.get(realFieldName)).mapToDouble(i -> (double) i).sum();
                            sheetContext.getReplaceMap().put(replaceName, sum);
                        }
                    }
                }
                break;
            case FORMULA:
                for (String replaceName : replaceNames) {
                    if (replaceName.endsWith(FORMULA_SUFFIX)) { // 公式 支持简单的 加减乘除
                        // 先获取真实的公式 ，然后再将字段替换为真实值   再计算 ， 并且公式的值必须在sheetContext map中能取到，所以需要最后执行
                        final String formula = replaceName.replaceAll(FORMULA_SUFFIX, "");
                        final String realFormula = FormulaUtils.convert2Formula(formula, sheetContext.getReplaceMap());
                        final Object o = FormulaUtils.analysisFormula(realFormula);
                        sheetContext.getReplaceMap().put(replaceName, o);
                    }
                }
                break;
            default:
                break;
        }
    }


    private static void setPageSize(Sheet sourceSheet, Sheet targetSheet) {

        // 拷贝页面设置
        PrintSetup srcPrintSetup = sourceSheet.getPrintSetup();
        PrintSetup destPrintSetup = targetSheet.getPrintSetup();
        destPrintSetup.setPaperSize(srcPrintSetup.getPaperSize());
        destPrintSetup.setFitHeight(srcPrintSetup.getFitHeight());
        destPrintSetup.setFitWidth(srcPrintSetup.getFitWidth());
        destPrintSetup.setLeftToRight(srcPrintSetup.getLeftToRight());
        destPrintSetup.setLandscape(srcPrintSetup.getLandscape());
        destPrintSetup.setHeaderMargin(srcPrintSetup.getHeaderMargin());
        destPrintSetup.setFooterMargin(srcPrintSetup.getFooterMargin());
        destPrintSetup.setCopies(srcPrintSetup.getCopies());
        destPrintSetup.setScale(srcPrintSetup.getScale());
        destPrintSetup.setPageStart(srcPrintSetup.getPageStart());
        destPrintSetup.setValidSettings(srcPrintSetup.getValidSettings());
        destPrintSetup.setLeftToRight(srcPrintSetup.getLeftToRight());

        targetSheet.setMargin(PageMargin.BOTTOM, sourceSheet.getMargin(PageMargin.BOTTOM));
        targetSheet.setMargin(PageMargin.LEFT, sourceSheet.getMargin(PageMargin.LEFT));
        targetSheet.setMargin(PageMargin.TOP, sourceSheet.getMargin(PageMargin.TOP));
        targetSheet.setMargin(PageMargin.RIGHT, sourceSheet.getMargin(PageMargin.RIGHT));
        targetSheet.setMargin(PageMargin.HEADER, sourceSheet.getMargin(PageMargin.HEADER));
        targetSheet.setMargin(PageMargin.FOOTER, sourceSheet.getMargin(PageMargin.FOOTER));
    }


    private static void copyTemplateType(Sheet sheetInput, Sheet outputSheet, Context context, Workbook workbookInput, Workbook outputWorkbook, String typeName, SheetContext sheetContext) {
        ExcelUtils.copyMergedCellStyles(sheetInput, outputSheet, context.getExcelTemplate().getTypeRow() - 1, outputSheet.getLastRowNum() + 1, 1);
        final Row typeRow = sheetInput.getRow(context.getExcelTemplate().getTypeRow() - 1);
        Row typeRowOutput = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
        for (int j = 0; j < typeRow.getLastCellNum(); j++) { // 遍历每一列
            Cell cellInput = typeRow.getCell(j);
            Cell cellOutput = typeRowOutput.createCell(j);
            if (cellInput != null) {
                ExcelUtils.setCellValue(cellInput, cellOutput, sheetContext);
                // todo jcj替换变量
                cellOutput.setCellStyle(ExcelUtils.getCopyCellStyle(cellInput.getCellStyle(), workbookInput, outputWorkbook, false));
                setCellWidth(sheetInput, outputSheet, j);
            }
        }
    }

    private static void copyTemplateHeader(Sheet sheetInput, Sheet outputSheet, Context context, Workbook workbookInput, Workbook outputWorkbook, SheetContext sheetContext) {
        int index = outputSheet.getLastRowNum() + 1;
        for (int i = 0; i < context.getExcelTemplate().getPreRows(); i++) {
            Row rowInput = sheetInput.getRow(i);
            Row rowOutput = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
            for (int j = 0; j < rowInput.getLastCellNum(); j++) { // 遍历每一列
                Cell cellInput = rowInput.getCell(j);
                Cell cellOutput = rowOutput.createCell(j);
                if (cellInput != null) {
                    setCellWidth(sheetInput, outputSheet, j);
                    ExcelUtils.setCellValue(cellInput, cellOutput, sheetContext);
                    cellOutput.setCellStyle(ExcelUtils.getCopyCellStyle(cellInput.getCellStyle(), workbookInput, outputWorkbook, true));
                }
            }
        }
        ExcelUtils.copyMergedCellStyles(sheetInput, outputSheet, 0, index, context.getExcelTemplate().getPreRows());
    }

    private static void setCellWidth(Sheet srcSheet, Sheet destSheet, int index) {
        destSheet.setColumnWidth(index, srcSheet.getColumnWidth(index));
    }


    private static void copyTemplateBottom(Sheet sheetInput, Sheet outputSheet, Context context, Workbook workbookInput, Workbook outputWorkbook, SheetContext sheetContext) {

        int startRow = ExcelUtils.getLastRowWithData(sheetInput) - context.getExcelTemplate().getLastRows() + 1;
        ExcelUtils.copyMergedCellStyles(sheetInput, outputSheet, startRow, outputSheet.getLastRowNum() + 1, context.getExcelTemplate().getLastRows());
        for (int i = 0; i < context.getExcelTemplate().getLastRows(); i++, startRow++) {
            Row rowInput = sheetInput.getRow(startRow);
            Row rowOutput = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
            rowOutput.setRowStyle(rowInput.getRowStyle());
            for (int j = 0; j < rowInput.getLastCellNum(); j++) { // 遍历每一列
                Cell cellInput = rowInput.getCell(j);
                Cell cellOutput = rowOutput.createCell(j);
                if (cellInput != null) {
                    ExcelUtils.setCellValue(cellInput, cellOutput, sheetContext);
                    cellOutput.setCellStyle(ExcelUtils.getCopyCellStyle(cellInput.getCellStyle(), workbookInput, outputWorkbook, false));
                    setCellWidth(sheetInput, outputSheet, j);
                }
            }
        }

    }


    public static void writeExcel(String fileName, List<Map<String, Object>> data, Pair<List<Cell>, List<Cell>> preLastCell, Context context) throws IOException {
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Sheet1");
        // 写入前N行cell

        Workbook workbook = new XSSFWorkbook(new FileInputStream(context.getExcelTemplate().getTemplateFile()));
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
                case FORMULA:
                    if (currentCell.getCachedFormulaResultType() == CellType.NUMERIC) { // 数字公式 直接返回显示结果，不返回公式
                        cellValue = currentCell.getNumericCellValue();
                    } else {
                        cellValue = currentCell.getStringCellValue();
                    }
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(currentCell)) {
                        final SimpleDateFormat simpleDateFormat = SimpleDateThreadLocal.getSimpleDateFormat("yyyy-MM-dd");
                        cellValue = simpleDateFormat.format(currentCell.getDateCellValue());
                    } else {
                        cellValue = currentCell.getNumericCellValue();
                    }
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
        final ArrayList<String> tempStr = Lists.newArrayList();
        // 遍历数据列表
        for (Map<String, Object> item : data) {
            // 根据分类字段的值获取对应的键
            for (String allocationField : allocation) {
                tempStr.add((String) item.get(allocationField));
            }

            final String allocationKey = String.join("-", tempStr);

            // 根据分类键将数据放入对应的列表中
            List<Map<String, Object>> group = groupedData.getOrDefault(allocationKey, new ArrayList<>());
            group.add(item);
            groupedData.put(allocationKey, group);
            tempStr.clear();
        }
        // 将分类结果转换为列表形式
        return groupedData;
    }


    public static Pair<List<Row>, List<Row>> extractRowsFromExcel(File templateFile, int preRows, int lastRows) throws IOException {
        List<Row> preRow = new ArrayList<>();
        List<Row> lastRow = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(templateFile));
        Sheet sheet = workbook.getSheetAt(0);
        int totalRows = sheet.getLastRowNum() + 1;
        // 抽取前N行单元格
        for (int i = 0; i < Math.min(preRows, totalRows); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                preRow.add(row);
            }
        }
        // 抽取最后X行单元格
        for (int i = Math.max(totalRows - lastRows, 0); i < totalRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                lastRow.add(row);
            }
        }

        return Pair.create(preRow, lastRow);
    }
}
