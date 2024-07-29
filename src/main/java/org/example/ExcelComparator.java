package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelComparator {

    private static final String PARTIAL_HEADER = "Код позиции";
    private static final Logger logger = LoggerFactory.getLogger(ExcelComparator.class);
    private static final Map<String, CellStyle> styleMap = new HashMap<>();

    public static void main(String[] args) {
        String file1 = "Выгрузка_МТР_ГСП_Ремонт__Подрядчик__24.07.2024.xls";
        String file2 = "Выгрузка_ООО_ГТТ_2024_13.06.2024_форма.xls";
        String outputFile = "Сравнение_результатов.xlsx";

        try {
            compareExcelFiles(file1, file2, outputFile);
        } catch (IOException e) {
            logger.error("Ошибка при сравнении Excel файлов", e);
        }
    }

    private static void compareExcelFiles(String file1, String file2, String outputFile) throws IOException {
        try (Workbook workbook1 = openWorkbook(file1);
             Workbook workbook2 = openWorkbook(file2);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet sheet1 = workbook1.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(0);

            // Создание листов для результатов
            Sheet headerSheet = outputWorkbook.createSheet("Заголовки");
            Sheet missingRowsSheet = outputWorkbook.createSheet("Отсутствующие строки");
            Sheet changedRowsSheet = outputWorkbook.createSheet("Измененные строки");

            // Сравнение заголовков
            Row header1 = sheet1.getRow(1);
            Row header2 = sheet2.getRow(1);
            boolean headersChanged = !compareHeaders(header1, header2, headerSheet);
            if (headersChanged) {
                createOutputRow(headerSheet, 1, "Заголовки изменились");
            }

            // Получение индексов колонок
            int index1 = getColumnIndexByPartialHeader(header1, PARTIAL_HEADER);
            int index2 = getColumnIndexByPartialHeader(header2, PARTIAL_HEADER);
            if (index1 == -1 || index2 == -1) {
                throw new RuntimeException("Заголовок, содержащий '" + PARTIAL_HEADER + "', не найден в одном из файлов.");
            }

            // Сравнение строк
            compareRows(sheet1, sheet2, missingRowsSheet, changedRowsSheet, index1, index2);

            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fileOut);
            }
            logger.info("Результаты сравнения сохранены в файл: " + outputFile);
        }
    }

    private static Workbook openWorkbook(String filePath) throws IOException {
        if (filePath.endsWith(".xlsx")) {
            return new XSSFWorkbook(new FileInputStream(filePath));
        } else if (filePath.endsWith(".xls")) {
            return new HSSFWorkbook(new FileInputStream(filePath));
        } else {
            throw new IllegalArgumentException("Неподдерживаемый формат файла: " + filePath);
        }
    }

    private static void compareRows(Sheet sheet1, Sheet sheet2, Sheet missingRowsSheet, Sheet changedRowsSheet, int index1, int index2) {
        int missingRowNum = 1;
        int changedRowNum = 1;

        // Создание отображений значений строк
        Map<String, Row> rowMap1 = createValueToRowMap(sheet1, index1);
        Map<String, Row> rowMap2 = createValueToRowMap(sheet2, index2);

        // Поиск строк, отсутствующих во втором файле
        for (Map.Entry<String, Row> entry : rowMap1.entrySet()) {
            String value = entry.getKey();
            Row row1 = entry.getValue();
            Row matchingRow = rowMap2.get(value);

            if (matchingRow == null) {
                createOutputRow(missingRowsSheet, missingRowNum++, "Строка с '" + PARTIAL_HEADER + "' = " + value + " отсутствует во втором файле.");
            } else if (!compareRows(row1, matchingRow)) {
                createOutputRow(changedRowsSheet, changedRowNum++, "Строка с '" + PARTIAL_HEADER + "' = " + value + " изменилась:");
                changedRowNum = printRowDifferences(row1, matchingRow, changedRowsSheet, changedRowNum);
            }
        }
    }

    private static Map<String, Row> createValueToRowMap(Sheet sheet, int columnIndex) {
        Map<String, Row> valueToRowMap = new HashMap<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String value = getCellValue(row.getCell(columnIndex));
            valueToRowMap.put(value, row);
        }
        return valueToRowMap;
    }

    private static void createOutputRow(Sheet sheet, int rowNum, String message) {
        Row row = sheet.createRow(rowNum);
        Cell cell = row.createCell(0);
        cell.setCellValue(message);

        CellStyle style = getCellStyle(sheet.getWorkbook(), "Bold");
        cell.setCellStyle(style);
    }

    private static CellStyle getCellStyle(Workbook workbook, String styleKey) {
        if (!styleMap.containsKey(styleKey)) {
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            styleMap.put(styleKey, style);
        }
        return styleMap.get(styleKey);
    }

    private static boolean compareHeaders(Row header1, Row header2, Sheet headerSheet) {
        if (header1 == null || header2 == null) {
            createOutputRow(headerSheet, 1, "Один из заголовков отсутствует.");
            return false;
        }

        Map<String, Integer> headerMap1 = createHeaderMap(header1);
        Map<String, Integer> headerMap2 = createHeaderMap(header2);

        int rowNum = 0;
        for (Map.Entry<String, Integer> entry : headerMap1.entrySet()) {
            String header1Value = entry.getKey();
            Integer columnIndex2 = headerMap2.get(header1Value);

            if (columnIndex2 == null) {
                createOutputRow(headerSheet, ++rowNum, "Заголовок '" + header1Value + "' из первого файла отсутствует во втором.");
            } else {
                headerMap2.remove(header1Value);
            }
        }

        for (Map.Entry<String, Integer> entry : headerMap2.entrySet()) {
            String header2Value = entry.getKey();
            createOutputRow(headerSheet, ++rowNum, "Заголовок '" + header2Value + "' из второго файла отсутствует в первом.");
        }

        if (rowNum == 0) {
            createOutputRow(headerSheet, ++rowNum, "Заголовки совпадают.");
        }

        return rowNum > 1; // Return true if headers changed
    }

    private static Map<String, Integer> createHeaderMap(Row headerRow) {
        Map<String, Integer> headerMap = new HashMap<>();
        if (headerRow == null) return headerMap;

        for (Cell cell : headerRow) {
            if (cell != null && cell.getCellType() == CellType.STRING) {
                headerMap.put(cell.getStringCellValue(), cell.getColumnIndex());
            }
        }
        return headerMap;
    }

    private static boolean compareRows(Row row1, Row row2) {
        if (row1 == null || row2 == null) return row1 == row2;

        int maxCells = Math.max(row1.getLastCellNum(), row2.getLastCellNum());
        for (int i = 0; i < maxCells; i++) {
            Cell cell1 = row1.getCell(i);
            Cell cell2 = row2.getCell(i);

            if (!compareCells(cell1, cell2)) {
                return false;
            }
        }
        return true;
    }

    private static boolean compareCells(Cell cell1, Cell cell2) {
        if (cell1 == null && cell2 == null) return true;
        if (cell1 == null || cell2 == null) return false;

        if (cell1.getCellType() != cell2.getCellType()) return false;

        switch (cell1.getCellType()) {
            case STRING:
                return cell1.getStringCellValue().equals(cell2.getStringCellValue());
            case NUMERIC:
                return cell1.getNumericCellValue() == cell2.getNumericCellValue();
            case BOOLEAN:
                return cell1.getBooleanCellValue() == cell2.getBooleanCellValue();
            case FORMULA:
                return cell1.getCellFormula().equals(cell2.getCellFormula());
            default:
                return false;
        }
    }

    private static int printRowDifferences(Row row1, Row row2, Sheet sheet, int startRow) {
        // Создаем стили для вывода различий
        CellStyle redStyle = getCellStyle(sheet.getWorkbook(), "Red");

        for (int i = 0; i < Math.max(row1.getLastCellNum(), row2.getLastCellNum()); i++) {
            Cell cell1 = row1.getCell(i);
            Cell cell2 = row2.getCell(i);

            // Получаем значения ячеек
            String value1 = getCellValue(cell1);
            String value2 = getCellValue(cell2);

            // Сравниваем значения
            if (!value1.equals(value2)) {
                Row outputRow = sheet.createRow(startRow++);
                Cell outputCell = outputRow.createCell(0);
                outputCell.setCellValue("Различие в столбце " + i + ": " + value1 + " vs " + value2);
                outputCell.setCellStyle(redStyle);
            }
        }
        return startRow;
    }


    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static int getColumnIndexByPartialHeader(Row headerRow, String partialHeader) {
        if (headerRow == null) return -1;
        for (Cell cell : headerRow) {
            if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().contains(partialHeader)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }
}
