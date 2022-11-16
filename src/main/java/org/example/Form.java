package org.example;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.codec.Decoder;
import org.apache.commons.codec.binary.Base64OutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.io.*;
import java.lang.reflect.Array;
import java.net.URI;
import java.nio.Buffer;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;

public class Form {
//  元ファイルのコピー用suffix（このファイルに対して修正する）
    private static final String SUFFIX = "_temp.xlsx";
//  セルごとの項目群のキー / 明細部の項目群のキー
    private static final String FILED_KEY_CELL = "cell";
    private static final String FILED_KEY_ROW = "row";
    private static final ObjectMapper mapper = new ObjectMapper();
    private final JsonNode cellNode;
    private final JsonNode rowNode;
    private final Iterator<String> cellNames;
    private final Iterator<String> rowNames;
    private final XSSFWorkbook workbook;
    private final Sheet sheet;
    private final String filePath;
    public Form(String srcPath, String json) throws IOException {
        JsonNode _map = mapper.readTree(json);
        cellNode = _map.get(FILED_KEY_CELL);
        rowNode = _map.get(FILED_KEY_ROW);
        if (Objects.isNull(cellNode)) {
            cellNames = null;
        } else {
            cellNames = cellNode.fieldNames();
        }
        if (Objects.isNull(rowNode)) {
            rowNames = null;
        } else {
            rowNames = rowNode.fieldNames();
        }
        filePath = srcPath + SUFFIX;
        Files.copy(Paths.get(srcPath), Paths.get(filePath), StandardCopyOption.REPLACE_EXISTING);
        workbook = new XSSFWorkbook(filePath);
        sheet = workbook.getSheetAt(0);
    }
    public void setValues() {
        if (Objects.isNull(cellNames)) {
            return;
        }
        while(cellNames.hasNext()) {
            var cellName = cellNames.next();
            Name name = workbook.getName(cellName);
            // TODO nameがNULLの場合は例外とするか何もしないか。例外としてメッセージを返した方が親切？
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
        if (Objects.isNull(rowNames)) {
            return;
        }
        while(rowNames.hasNext()) {
            String rowName = rowNames.next();
            Name name = workbook.getName(rowName);
            // TODO nameがNULLの場合は例外とするか何もしないか。例外としてメッセージを返した方が親切？
            if (Objects.isNull(name)) {
                continue;
            }
            // 対象の名前セルの1件あたりの行数を取得
            CellRangeAddress nameRange = CellRangeAddress.valueOf(name.getRefersToFormula());
            int rowsPerItem = nameRange.getLastRow() - nameRange.getFirstRow() + 1;
            int startRowNum = nameRange.getLastRow() + 1;
            // JSONの行データ分
            JsonNode rows = rowNode.get(rowName).get("value");
            for (int i=0; i<rows.size(); i++) {
                // 基準行が存在するので、行コピーは1回分少なくする
                if (i < rows.size()-1) {
                    prepareRows(nameRange, rowsPerItem, startRowNum);
                }
                JsonNode row = rows.get(i);
                setValueByName(row, rowsPerItem, i);
            }
        }
    }
    private void prepareRows(CellRangeAddress nameRange, int rowsPerItem, int startRowNum) {
        // 1件あたりの行数分空行を作成し、スタイルをコピーしていく
        shiftRows(startRowNum, rowsPerItem);
        for (int i=0; i<rowsPerItem; i++) {
            Row sourceRow = sheet.getRow(nameRange.getFirstRow() + i);
            Row destinationRow = sheet.getRow(nameRange.getFirstRow() + rowsPerItem + i);
            copyStyle(sourceRow, destinationRow);
            // セル結合されている場合は、結合設定もコピー
            for (int j=0; j<sheet.getNumMergedRegions(); j++) {
                CellRangeAddress mergedRegion = sheet.getMergedRegion(j);
                if (mergedRegion.getFirstRow() != (nameRange.getFirstRow() + i)) continue;
                mergeCell(mergedRegion, rowsPerItem);
            }
        }
    }
    private void shiftRows(int startRowNum, int createRowCount) {
        for (int i=0; i<createRowCount; i++) {
            if (sheet.getRow(startRowNum + i) != null) {
                sheet.shiftRows(startRowNum + i, sheet.getLastRowNum(), 1);
            }
            sheet.createRow(startRowNum + i);
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
        Cell nextCell = cell;
        for (JsonNode jsonNode : nodeValue) {
            setValue(nextCell, jsonNode);
            nextCell = getCellAtNextRow(nextCell);
        }
    }
    private void setValue(Cell cell, JsonNode nodeValue) {
        cell.setCellValue(nodeValue.get("value").asText());
    }
    private void setValueByName(JsonNode row, int rowsPerItem, int index) {
        Iterator<String> rowFields = row.fieldNames();
        while(rowFields.hasNext()) {
            String fieldInRow = rowFields.next();
            Name cellName = workbook.getName(fieldInRow);
            if (Objects.isNull(cellName)) {
                continue;
            }
            CellRangeAddress range = CellRangeAddress.valueOf(cellName.getRefersToFormula());
            // 基準行から1行あたりの行数 * ループ回数分移動したセルを取得
            Cell topLeftCell = getTopLeftCell(range);
            int nextRowNum = topLeftCell.getRowIndex() + rowsPerItem * index;
//            System.out.println(nextRowNum);
            Cell nextCell = sheet.getRow(nextRowNum).getCell(topLeftCell.getColumnIndex());
            setValue(nextCell, row.get(fieldInRow));
        }
    }
    private void copyStyle(Row sourceRow, Row destinationRow) {
        for (int i=0; i<sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell destinationCell = destinationRow.createCell(i);
            if (sourceCell == null) continue;
            // cellStyle
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.cloneStyleFrom(sourceCell.getCellStyle());
            destinationCell.setCellStyle(cellStyle);
        }
    }
    private void mergeCell(CellRangeAddress sourceRange, int rowsPerItem) {
        CellRangeAddress mergeCellRangeAddress = new CellRangeAddress(
                sourceRange.getFirstRow() + rowsPerItem,
                sourceRange.getLastRow() + rowsPerItem,
                sourceRange.getFirstColumn(),
                sourceRange.getLastColumn()
        );
        sheet.addMergedRegion(mergeCellRangeAddress);
    }
    public String writeFile() throws IOException {
        ByteArrayOutputStream byteaOutput = new ByteArrayOutputStream();
        Base64OutputStream base64Output = new Base64OutputStream(byteaOutput);
        workbook.write(base64Output);
        workbook.close();
        byteaOutput.flush();
        byteaOutput.close();
        base64Output.flush();
        base64Output.close();
        Files.delete(Paths.get(filePath));
        return byteaOutput.toString();
    }
}
