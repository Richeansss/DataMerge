package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

public class ExcelMerger {
    private static final Logger logger = LoggerFactory.getLogger(ExcelMerger.class);
    private static final String FILE1_PATH = "ГТТ ДТОИР/ООРТОДО/Выгрузка_ООО_ГТТ_2024_13.06.2024_форма.xlsx";
    private static final String FILE2_PATH = "ГТТ ДТОИР/ООРТОДО/МТР_подрядчика_2024_26.06.2024xlsx.xlsx";
    private static final int KEY_COLUMN_FILE1 = 3;
    private static final int KEY_COLUMN_FILE2 = 6;
    private static final int DEFAULT_COLUMN_WIDTH = 15; // Adjust as needed

    public static void main(String[] args) {
        try {
            Workbook workbook1 = new XSSFWorkbook(new FileInputStream(FILE1_PATH));
            Workbook workbook2 = new XSSFWorkbook(new FileInputStream(FILE2_PATH));

            Map<String, Row> dataFile1 = extractData(workbook1, KEY_COLUMN_FILE1);
            Map<String, Row> dataFile2 = extractData(workbook2, KEY_COLUMN_FILE2);

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet mergedSheet = newWorkbook.createSheet("MergedData");

            // Create header row with source indication
            createHeaderRow(mergedSheet, workbook1, workbook2);

            applyColumnStyles(mergedSheet, 0, dataFile1.get(dataFile1.keySet().iterator().next()).getLastCellNum(), "file1");
            applyColumnStyles(mergedSheet, dataFile1.get(dataFile1.keySet().iterator().next()).getLastCellNum(), dataFile2.get(dataFile2.keySet().iterator().next()).getLastCellNum(), "file2");

            // Set column width
            for (int i = 0; i < mergedSheet.getRow(0).getLastCellNum(); i++) {
                mergedSheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH * 256); // 256 characters per unit width
            }

            int rowIndex = 1; // Start from the second row since the first row is for headers
            for (String key : dataFile1.keySet()) {
                if (dataFile2.containsKey(key)) {
                    Row row = mergedSheet.createRow(rowIndex++);
                    int cellIndex = 0;

                    // Copy data from file1
                    Row dataRow1 = dataFile1.get(key);
                    copyRowData(dataRow1, row, cellIndex, FILE1_PATH);

                    // Increment cellIndex by number of columns in dataRow1
                    cellIndex += dataRow1.getLastCellNum();

                    // Copy data from file2
                    Row dataRow2 = dataFile2.get(key);
                    copyRowData(dataRow2, row, cellIndex, FILE2_PATH);
                } else {
                    // If key is only in dataFile1, add to unmatched sheet 1
                    addUnmatchedRow(newWorkbook, dataFile1.get(key), "UnmatchedDataFromFile1");
                }
            }

            // Add unmatched rows from dataFile2 to unmatched sheet 2
            for (String key : dataFile2.keySet()) {
                if (!dataFile1.containsKey(key)) {
                    addUnmatchedRow(newWorkbook, dataFile2.get(key), "UnmatchedDataFromFile2");
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream("MergedData.xlsx")) {
                newWorkbook.write(fileOut);
            }


            workbook1.close();
            workbook2.close();
            TableColumnSorter.sortColumnsByHeaders(newWorkbook, "MergedData");
            newWorkbook.close();

        } catch (IOException e) {
            logger.error("Error processing Excel files", e);
        }
    }

    private static Map<String, Row> extractData(Workbook workbook, int keyColumnIndex) {
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

    private static void createHeaderRow(Sheet sheet, Workbook workbook1, Workbook workbook2) {
        Row headerRow = sheet.createRow(0);

        // Extract headers from workbook1 with source indication
        Sheet sheet1 = workbook1.getSheetAt(0);
        Row headerRow1 = sheet1.getRow(0);
        int cellIndex = 0;
        for (Cell cell : headerRow1) {
            Cell newCell = headerRow.createCell(cellIndex++);
            newCell.setCellValue(getCellValueAsString(cell) + " (from " + "file1" + ")");
        }

        // Extract headers from workbook2 with source indication
        Sheet sheet2 = workbook2.getSheetAt(0);
        Row headerRow2 = sheet2.getRow(0);
        for (Cell cell : headerRow2) {
            Cell newCell = headerRow.createCell(cellIndex++);
            newCell.setCellValue(getCellValueAsString(cell) + " (from " + "file2" + ")");
        }
    }

    private static void copyRowData(Row sourceRow, Row targetRow, int startIndex, String sourceFile) {
        if (sourceRow != null && targetRow != null) {
            int numCellsToCopy = sourceRow.getLastCellNum();
            for (int i = 0; i < numCellsToCopy; i++) {
                Cell sourceCell = sourceRow.getCell(i);
                Cell targetCell = targetRow.createCell(startIndex++);
                setCellValue(sourceCell, targetCell);
            }
        }
    }

    private static void setCellValue(Cell sourceCell, Cell targetCell) {
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


    private static String getCellValueAsString(Cell cell) {
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

    private static void applyColumnStyles(Sheet sheet, int startIndex, int numColumns, String source) {
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

            Row headerRow = sheet.getRow(0);
            Cell headerCell = headerRow.getCell(i);
            headerCell.setCellValue(headerCell.getStringCellValue() + " (from " + source + ")");
        }
    }



    private static void addUnmatchedRow(Workbook workbook, Row sourceRow, String sheetName) {
        Sheet unmatchedSheet = workbook.getSheet(sheetName);
        if (unmatchedSheet == null) {
            unmatchedSheet = workbook.createSheet(sheetName);

            // Create header row for unmatched sheet
            Row headerRow = unmatchedSheet.createRow(0);
            Sheet dummySheet = workbook.getSheetAt(0); // Use any existing sheet to get headers
            Row dummyHeaderRow = dummySheet.getRow(0);
            int cellIndex = 0;
            for (Cell cell : dummyHeaderRow) {
                Cell newCell = headerRow.createCell(cellIndex++);
                newCell.setCellValue(getCellValueAsString(cell));
            }
        }

        int rowIndex = unmatchedSheet.getLastRowNum() + 1;
        Row newRow = unmatchedSheet.createRow(rowIndex);
        copyRowData(sourceRow, newRow, 0, ""); // Source file path not needed for unmatched rows
    }

}
