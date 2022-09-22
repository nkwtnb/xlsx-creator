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
    private static final String TEMPLATE_PATH = "./resources/template2.xlsx";
    private static final String OUT_PATH = "./resources/out.xlsx";
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
            CellRangeAddress nameRange = CellRangeAddress.valueOf(name.getRefersToFormula());
            for (int i=0; i<sheet.getNumMergedRegions(); i++) {
                CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                if (mergedRegion.getFirstRow() != nameRange.getFirstRow()) continue;
                //
                sheet.getRow(mergedRegion.getLastRow()+1)
                copyRow(mergedRegion);
            }
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
    private void copyRow(CellRangeAddress sourceRange) {
        int sourceRowNum = sourceRange.getFirstRow();
        int mergedRowCount = (sourceRange.getLastRow() - sourceRange.getFirstRow()) + 1;
        System.out.println("range: " + sourceRange + "mergedRow: " + mergedRowCount);
        if (sheet.getRow(sourceRange.getLastRow()+1) == null) {
            System.out.println("destRow is null : " + (sourceRange.getLastRow()+1));
            sheet.createRow(sourceRange.getLastRow()+1);
        } else {
            System.out.println("destRow is not null : " + (sourceRange.getLastRow()+1));
            sheet.shiftRows(sourceRange.getLastRow()+1, sheet.getLastRowNum(), mergedRowCount);
        }
        Row destinationRow = sheet.getRow(sourceRange.getLastRow()+1);
        System.out.println("destRow: " + destinationRow);
        copyStyle(sourceRange, destinationRow);
        mergeCellAtSpecificRow(sourceRange, destinationRow);
    }
    private void copyStyle(CellRangeAddress sourceRange, Row destinationRow) {
        Row sourceRow = sheet.getRow(sourceRange.getFirstRow());
//        Row destinationRow = sheet.getRow(sourceRowNum+1);
        System.out.println("src: " + sourceRow.getRowNum());
        System.out.println("dest: " + destinationRow.getRowNum());
        for (int i=0; i<sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell destinationCell = destinationRow.createCell(i);
            // TODO sourceCellがnullの場合

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
    private void mergeCellAtSpecificRow(CellRangeAddress sourceRange, Row destinationRow) {
        int mergedRowCount = (sourceRange.getLastRow() - sourceRange.getFirstRow() + 1);
        System.out.println("mergedRegion: " + sourceRange);
//            System.out.println("param: " +
//                    destinationRow.getRowNum() + ", " +
//                    (destinationRow.getRowNum() + mergedRowCount) + ", " +
//                    mergedRegion.getFirstColumn() + ", " +
//                    mergedRegion.getLastColumn()
//            );
        CellRangeAddress mergeCellRangeAddress = new CellRangeAddress(
                destinationRow.getRowNum(),
                destinationRow.getRowNum() + mergedRowCount,
                sourceRange.getFirstColumn(),
                sourceRange.getLastColumn()
        );
        sheet.addMergedRegion(mergeCellRangeAddress);
//
//        for (int i=0; i<sheet.getNumMergedRegions(); i++) {
//            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
//            if (mergedRegion.getFirstRow() != sourceRange.getFirstRow()) continue;
//            int mergedRowCount = mergedRegion.getLastRow() - mergedRegion.getFirstRow();
//            System.out.println("mergedRegion: " + mergedRegion);
////            System.out.println("param: " +
////                    destinationRow.getRowNum() + ", " +
////                    (destinationRow.getRowNum() + mergedRowCount) + ", " +
////                    mergedRegion.getFirstColumn() + ", " +
////                    mergedRegion.getLastColumn()
////            );
//            CellRangeAddress mergeCellRangeAddress = new CellRangeAddress(
//                    destinationRow.getRowNum(),
//                    destinationRow.getRowNum() + mergedRowCount,
//                    mergedRegion.getFirstColumn(),
//                    mergedRegion.getLastColumn()
//            );
//            sheet.addMergedRegion(mergeCellRangeAddress);
//        }
    }
    public void writeFile() throws IOException {
        try (FileOutputStream os = new FileOutputStream(OUT_PATH)) {
            workbook.write(os);
        }
    }
}
