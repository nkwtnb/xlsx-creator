package org.example;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.Iterator;
import java.util.Objects;

public class Form {
    private static final String TEMPLATE_PATH = "./src/main/resources/template_multi.xlsx";
    private static final String OUT_PATH = "./src/main/resources/out.xlsx";
    private static final ObjectMapper mapper = new ObjectMapper();
    private final JsonNode cellNode;
    private final JsonNode rowNode;
    private final Iterator<String> cellNames;
    private final Iterator<String> rowNames;
    private final Workbook workbook;
    private final Sheet sheet;
    public Form(String json) throws IOException {
        JsonNode _map = mapper.readTree(json);
        cellNode = _map.get("cell");
        rowNode = _map.get("row");
        // TODO null check
        cellNames = cellNode.fieldNames();
        rowNames = rowNode.fieldNames();
        workbook = WorkbookFactory.create(new File(TEMPLATE_PATH));
        sheet = workbook.getSheetAt(0);
    }
    public void setValues() {
        while(cellNames.hasNext()) {
            var cellName = cellNames.next();
            Name name = workbook.getName(cellName);
            // TODO nameがNULLの場合はExceptionとするか何もしないか。とりあえず何もしない。
            if (Objects.isNull(name)) {
                continue;
            }
            CellRangeAddress range = CellRangeAddress.valueOf(name.getRefersToFormula());
            Cell cell = getTopLeftCell(range);
            JsonNode value = cellNode.get(cellName);
            if (value.isArray()) {
                iterateArrayValue(cell, value);
            } else {
                setValue(cell, value);
            }
        }
    }
    public void setRowValues() {
        while(rowNames.hasNext()) {
            String rowName = rowNames.next();
            Name name = workbook.getName(rowName);
            // TODO nameがNULLの場合はExceptionとするか何もしないか。とりあえず何もしない。
            if (Objects.isNull(name)) {
                continue;
            }
            // 対象の名前セルの1件あたりの行数
            CellRangeAddress nameRange = CellRangeAddress.valueOf(name.getRefersToFormula());
            int rowsPerItem = nameRange.getLastRow() - nameRange.getFirstRow() + 1;
            prepareNextRow(nameRange, rowsPerItem);
            for (int i=0; i<rowsPerItem; i++) {
                Row sourceRow = sheet.getRow(nameRange.getFirstRow() + i);
                Row destinationRow = sheet.getRow(nameRange.getFirstRow() + rowsPerItem + i);
                copyRow(sourceRow, destinationRow);
                for (int j=0; j<sheet.getNumMergedRegions(); j++) {
                    CellRangeAddress mergedRegion = sheet.getMergedRegion(j);
                    System.out.println(mergedRegion);
                    if (mergedRegion.getFirstRow() != (nameRange.getFirstRow() + i)) continue;
                    mergeCellAtSpecificRow(mergedRegion, rowsPerItem);
                }
            }
        }
    }
    private void prepareNextRow(CellRangeAddress range, int rowsPerItem) {
        // コピー元行の次の行から
        int rowNum = range.getLastRow() + 1;
        // 1件あたりの行数分シフト
        for (int i=0; i<rowsPerItem; i++) {
            if (sheet.getRow(rowNum + i) != null) {
                sheet.shiftRows(rowNum + i, sheet.getLastRowNum(), 1);
            }
            sheet.createRow(rowNum + i);
        }
    }
    private Cell getTopLeftCell(CellRangeAddress range) {
        Row row = sheet.getRow(range.getFirstRow());
        return row.getCell(range.getFirstColumn());
    }
    private Cell getCellAtNextRow(Cell cell) {
        int currentRowNum = cell.getRowIndex();
        Row nextRow = sheet.getRow(currentRowNum+1);
        return nextRow.getCell(cell.getColumnIndex());
    }
    private void iterateArrayValue(Cell cell, JsonNode nodeValue) {
        System.out.println("This is iterateArray");
        Cell nextCell = cell;
        for (JsonNode jsonNode : nodeValue) {
            setValue(nextCell, jsonNode);
            nextCell = getCellAtNextRow(nextCell);
        }
    }
    private void setValue(Cell cell, JsonNode nodeValue) {
        System.out.println("This is String value : " + nodeValue);
        cell.setCellValue(nodeValue.asText());
    }
    private void copyRow(Row sourceRow, Row destinationRow) {
        for (int i=0; i<sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell destinationCell = destinationRow.createCell(i);
            if (sourceCell == null) continue;
            // cellStyle
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.cloneStyleFrom(sourceCell.getCellStyle());
            destinationCell.setCellStyle(cellStyle);

            // cellValue
            switch(sourceCell.getCellType()) {
                case NUMERIC -> {
                    destinationCell.setCellValue(sourceCell.getNumericCellValue());
                }
                case STRING -> {
                    destinationCell.setCellValue(sourceCell.getStringCellValue());
                }
            }
        }
    }
    private void mergeCellAtSpecificRow(CellRangeAddress sourceRange, int rowsPerItem) {
        CellRangeAddress mergeCellRangeAddress = new CellRangeAddress(
                sourceRange.getFirstRow() + rowsPerItem,
                sourceRange.getLastRow() + rowsPerItem,
                sourceRange.getFirstColumn(),
                sourceRange.getLastColumn()
        );
        sheet.addMergedRegion(mergeCellRangeAddress);
    }
    public void writeFile() throws IOException {
        try (FileOutputStream os = new FileOutputStream(OUT_PATH)) {
            workbook.write(os);
        }
    }
}
