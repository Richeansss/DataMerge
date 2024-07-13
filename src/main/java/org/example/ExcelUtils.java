package org.example;

import org.apache.poi.ss.usermodel.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class ExcelUtils {

    public static Map<String, Row> extractData(Workbook workbook, int keyColumnIndex) {
        Map<String, Row> dataMap = new HashMap<>();
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            Cell cell = row.getCell(keyColumnIndex);
            if (cell != null) {
                String key = getCellValueAsString(cell);
                dataMap.put(key, row);
            }
        }
        return dataMap;
    }

    public static void createHeaderRow(Sheet sheet, Workbook workbook1, Workbook workbook2) {
        Row headerRow = sheet.createRow(0);
        Set<String> addedHeaders = new HashSet<>();

        Sheet sheet1 = workbook1.getSheetAt(0);
        Row headerRow1 = sheet1.getRow(0);
        int cellIndex = 0;
        for (Cell cell : headerRow1) {
            String headerValue = getCellValueAsString(cell);
            String newHeaderValue = headerValue + " (from file1)";
            if (addedHeaders.contains(newHeaderValue)) {
                newHeaderValue = headerValue + " (from file1, duplicate)";
            }
            Cell newCell = headerRow.createCell(cellIndex++);
            newCell.setCellValue(newHeaderValue);
            addedHeaders.add(newHeaderValue);
        }

        Sheet sheet2 = workbook2.getSheetAt(0);
        Row headerRow2 = sheet2.getRow(0);
        for (Cell cell : headerRow2) {
            String headerValue = getCellValueAsString(cell);
            String newHeaderValue = headerValue + " (from file2)";
            if (addedHeaders.contains(newHeaderValue)) {
                newHeaderValue = headerValue + " (from file2, duplicate)";
            }
            Cell newCell = headerRow.createCell(cellIndex++);
            newCell.setCellValue(newHeaderValue);
            addedHeaders.add(newHeaderValue);
        }
    }

    public static void copyRowData(Row sourceRow, Row targetRow, int startIndex, String sourceFile) {
        if (sourceRow != null && targetRow != null) {
            int numCellsToCopy = sourceRow.getLastCellNum();
            for (int i = 0; i < numCellsToCopy; i++) {
                Cell sourceCell = sourceRow.getCell(i);
                Cell targetCell = targetRow.createCell(startIndex++);
                setCellValue(sourceCell, targetCell);
            }
        }
    }

    public static void setCellValue(Cell sourceCell, Cell targetCell) {
        if (sourceCell != null) {
            switch (sourceCell.getCellType()) {
                case STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(sourceCell)) {
                        targetCell.setCellValue(sourceCell.getDateCellValue());
                    } else {
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    targetCell.setCellValue(sourceCell.getCellFormula());
                    break;
                case BLANK:
                    targetCell.setCellValue(""); // Handle empty cells
                    break;
                default:
                    targetCell.setCellValue(sourceCell.toString());
            }
        } else {
            targetCell.setCellValue(""); // Handle null cells
        }
    }

    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
                    return df.format(cell.getDateCellValue());
                } else {
                    return Double.toString(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return cell.toString();
        }
    }

    public static void applyColumnStyles(Sheet sheet, int startIndex, int numColumns, String source) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        for (int i = startIndex; i < startIndex + numColumns; i++) {
            CellStyle columnStyle = workbook.createCellStyle();
            columnStyle.cloneStyleFrom(style);

            // Set different colors based on source
            if (source.equals("file1")) {
                columnStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
            } else if (source.equals("file2")) {
                columnStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            } else {
                columnStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            }

            columnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            sheet.setDefaultColumnStyle(i, columnStyle);
        }
    }

    public static void addUnmatchedRow(Workbook workbook, Row sourceRow, String sheetName) {
        Sheet unmatchedSheet = workbook.getSheet(sheetName);
        if (unmatchedSheet == null) {
            unmatchedSheet = workbook.createSheet(sheetName);

            // Create header row for unmatched sheet based on sourceRow headers
            Row headerRow = unmatchedSheet.createRow(0);
            for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                Cell sourceCell = sourceRow.getCell(i);
                Cell newCell = headerRow.createCell(i);
                newCell.setCellValue(getCellValueAsString(sourceCell));
            }
        } else {
            // Check if header row exists, create if not
            if (unmatchedSheet.getRow(0) == null) {
                Row headerRow = unmatchedSheet.createRow(0);
                for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                    Cell sourceCell = sourceRow.getCell(i);
                    Cell newCell = headerRow.createCell(i);
                    newCell.setCellValue(getCellValueAsString(sourceCell));
                }
            }
        }

        int rowIndex = unmatchedSheet.getLastRowNum() + 1;
        Row newRow = unmatchedSheet.createRow(rowIndex);
        copyRowData(sourceRow, newRow, 0, ""); // Source file path not needed for unmatched rows
    }
}
