package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import static org.example.GroupRowsByPositionAndCount.groupRowsAndCountInSameFile;
import static org.example.XlsToXlsxConverter.convertXlsToXlsx;

/**
 * Класс для объединения данных из двух Excel файлов.
 */
public class ExcelMerger {
    private static final Logger logger = LoggerFactory.getLogger(ExcelMerger.class);
    private static final String FILE1_PATH = "ДТОиР студентам/Выгрузка_ООО_ГТТ_2024_17.07.2024_форма.xlsx";
    private static final String FILE2_PATH = "ДТОиР студентам/МТР_подрядчика_2024_17__08.07.2024.xlsx";
    private static final int KEY_COLUMN_FILE1 = 3;
    private static final int KEY_COLUMN_FILE2 = 7;
    private static final int DEFAULT_COLUMN_WIDTH = 20; // Ширина строки

    /**
     * Основной метод для объединения данных из двух Excel файлов.
     *
     * @param args аргументы командной строки (не используются)
     * @throws IOException если возникают ошибки ввода-вывода при обработке файлов
     */
    public static void main(String[] args) throws IOException {
        try {
            // Проверка и конвертация файлов, если они в формате XLS
            String convertedFile1Path = convertIfNecessary(FILE1_PATH);
            String convertedFile2Path = convertIfNecessary(FILE2_PATH);

            // Открытие рабочих книг для файлов
            Workbook workbook1 = new XSSFWorkbook(new FileInputStream(convertedFile1Path));
            Workbook workbook2 = new XSSFWorkbook(new FileInputStream(convertedFile2Path));

            // Извлечение данных из файлов в виде карты ключ-строка
            Map<String, Row> dataFile1 = ExcelUtils.extractData(workbook1, KEY_COLUMN_FILE1);
            Map<String, Row> dataFile2 = ExcelUtils.extractData(workbook2, KEY_COLUMN_FILE2);

            // Создание новой рабочей книги для объединенных данных
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet mergedSheet = newWorkbook.createSheet("MergedData");

            // Создание заголовков с указанием источника для объединенного листа
            ExcelUtils.createHeaderRow(mergedSheet, workbook1, workbook2);

            // Применение стилей столбцов в зависимости от источника данных
            ExcelUtils.applyColumnStyles(mergedSheet, 0, dataFile1.get(dataFile1.keySet().iterator().next()).getLastCellNum(), "file1");
            ExcelUtils.applyColumnStyles(mergedSheet, dataFile1.get(dataFile1.keySet().iterator().next()).getLastCellNum(), dataFile2.get(dataFile2.keySet().iterator().next()).getLastCellNum(), "file2");

            // Установка ширины столбцов
            for (int i = 0; i < mergedSheet.getRow(0).getLastCellNum(); i++) {
                mergedSheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH * 256); // 256 символов на единицу ширины
            }

            int rowIndex = 1; // Начинаем с второй строки, так как первая строка для заголовков
            // Обработка данных из file1
            for (String key : dataFile1.keySet()) {
                if (dataFile2.containsKey(key)) {
                    Row row = mergedSheet.createRow(rowIndex++);
                    int cellIndex = 0;

                    // Копирование данных из file1
                    Row dataRow1 = dataFile1.get(key);
                    ExcelUtils.copyRowData(dataRow1, row, cellIndex, FILE1_PATH);

                    // Увеличение cellIndex на количество столбцов в dataRow1
                    cellIndex += dataRow1.getLastCellNum();

                    // Копирование данных из file2
                    Row dataRow2 = dataFile2.get(key);
                    ExcelUtils.copyRowData(dataRow2, row, cellIndex, FILE2_PATH);
                } else {
                    // Если ключ только в dataFile1, добавляем в непринятый лист 1
                    ExcelUtils.addUnmatchedRow(newWorkbook, dataFile1.get(key), "UnmatchedDataFromFile1");
                }
            }

            // Добавление непринятых строк из dataFile2 в непринятый лист 2
            for (String key : dataFile2.keySet()) {
                if (!dataFile1.containsKey(key)) {
                    ExcelUtils.addUnmatchedRow(newWorkbook, dataFile2.get(key), "UnmatchedDataFromFile2");
                }
            }

            // Сохранение объединенных данных в новом Excel файле
            try (FileOutputStream fileOut = new FileOutputStream("MergedData.xlsx")) {
                newWorkbook.write(fileOut);
            }

            // Закрытие рабочих книг
            workbook1.close();
            workbook2.close();

            // Сортировка столбцов по заголовкам в объединенном файле
            TableColumnSorter.sortColumnsByHeaders(newWorkbook, "MergedData");

            // Закрытие новой рабочей книги
            newWorkbook.close();

        } catch (IOException e) {
            logger.error("Ошибка при обработке Excel файлов", e);
        }

        // Группировка строк и подсчет одинаковых строк в том же файле
        groupRowsAndCountInSameFile("MergedData.xlsx", "SortedData");
    }

    /**
     * Метод проверяет, является ли файл в формате XLS и при необходимости конвертирует его в формат XLSX.
     *
     * @param filePath путь к файлу
     * @return путь к конвертированному файлу XLSX или исходный путь, если файл уже в формате XLSX
     * @throws IOException если возникают ошибки ввода-вывода при работе с файлами
     */
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
