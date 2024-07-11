package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import static org.example.XlsToXlsxConverter.convertXlsToXlsx;

public class ExcelMerger {
    private static final Logger logger = LoggerFactory.getLogger(ExcelMerger.class);
    private static final String FILE1_PATH = "ДТОиР студентам/ООРТОДО/Выгрузка_ООО_ГТТ_2024_24.06.2024_форма.xls";
    private static final String FILE2_PATH = "ДТОиР студентам/ООРТОДО/МТР_подрядчика_2024_17__08.07.2024.xlsx";
    private static final int KEY_COLUMN_FILE1 = 3;
    private static final int KEY_COLUMN_FILE2 = 7;
    private static final int DEFAULT_COLUMN_WIDTH = 15; // Ширина строки

    public static void main(String[] args) {
        try {
            // Проверка и конвертация файлов, если они в формате XLS
            String convertedFile1Path = convertIfNecessary(FILE1_PATH);
            String convertedFile2Path = convertIfNecessary(FILE2_PATH);

            Workbook workbook1 = new XSSFWorkbook(new FileInputStream(convertedFile1Path));
            Workbook workbook2 = new XSSFWorkbook(new FileInputStream(convertedFile2Path));

            Map<String, Row> dataFile1 = ExcelUtils.extractData(workbook1, KEY_COLUMN_FILE1);
            Map<String, Row> dataFile2 = ExcelUtils.extractData(workbook2, KEY_COLUMN_FILE2);

            Workbook newWorkbook = new XSSFWorkbook();
            Sheet mergedSheet = newWorkbook.createSheet("MergedData");

            // Создание заголовков с указанием источника
            ExcelUtils.createHeaderRow(mergedSheet, workbook1, workbook2);

            ExcelUtils.applyColumnStyles(mergedSheet, 0, dataFile1.get(dataFile1.keySet().iterator().next()).getLastCellNum(), "file1");
            ExcelUtils.applyColumnStyles(mergedSheet, dataFile1.get(dataFile1.keySet().iterator().next()).getLastCellNum(), dataFile2.get(dataFile2.keySet().iterator().next()).getLastCellNum(), "file2");

            // Установка ширины столбца
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
                    ExcelUtils.copyRowData(dataRow1, row, cellIndex, FILE1_PATH);

                    // Increment cellIndex by number of columns in dataRow1
                    cellIndex += dataRow1.getLastCellNum();

                    // Copy data from file2
                    Row dataRow2 = dataFile2.get(key);
                    ExcelUtils.copyRowData(dataRow2, row, cellIndex, FILE2_PATH);
                } else {
                    // If key is only in dataFile1, add to unmatched sheet 1
                    ExcelUtils.addUnmatchedRow(newWorkbook, dataFile1.get(key), "UnmatchedDataFromFile1");
                }
            }

            // Add unmatched rows from dataFile2 to unmatched sheet 2
            for (String key : dataFile2.keySet()) {
                if (!dataFile1.containsKey(key)) {
                    ExcelUtils.addUnmatchedRow(newWorkbook, dataFile2.get(key), "UnmatchedDataFromFile2");
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
    private static String convertIfNecessary(String filePath) throws IOException {
        if (filePath.endsWith(".xls")) {
            logger.info("Обнаружен файл в формате XLS: {}", filePath);
            String xlsxFilePath = filePath.replace(".xls", ".xlsx");
            convertXlsToXlsx(filePath, xlsxFilePath);
            return xlsxFilePath;
        } else if (filePath.endsWith(".xlsx")) {
            return filePath;
        } else {
            throw new IOException("Неподдерживаемый формат файла: " + filePath);
        }
    }
}
