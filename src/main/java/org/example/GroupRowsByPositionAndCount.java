package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Класс для группировки строк в Excel-файле по позиции и подсчета количества строк в каждой группе.
 */
public class GroupRowsByPositionAndCount {
    private static final Logger logger = LoggerFactory.getLogger(GroupRowsByPositionAndCount.class);

    /**
     * Группирует строки и подсчитывает их количество в одном и том же файле.
     *
     * @param inputFilePath Путь к входному файлу Excel.
     * @param sheetName     Имя листа, на котором будет выполняться группировка.
     * @throws IOException Если возникнет ошибка при чтении или записи файла.
     */
    public static void groupRowsAndCountInSameFile(String inputFilePath, String sheetName) throws IOException {
        logger.info("Чтение файла: {}", inputFilePath);
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);

        groupAndCountRows(sheet);

        // Сохранение измененной рабочей книги обратно в тот же файл
        FileOutputStream fos = new FileOutputStream(inputFilePath);
        workbook.write(fos);
        fos.close();
        workbook.close();
        fis.close();
        logger.info("Файл успешно сохранен: {}", inputFilePath);
    }

    /**
     * Группирует строки на листе и подсчитывает их количество.
     *
     * @param sheet Лист, на котором будет выполняться группировка.
     */
    public static void groupAndCountRows(Sheet sheet) {
        int positionColumnIndex = -1;
        Row headerRow = sheet.getRow(0);

        // Поиск индекса колонки "ППП (from file2)"
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equals("ППП (from file2)")) {
                positionColumnIndex = cell.getColumnIndex();
                break;
            }
        }

        if (positionColumnIndex == -1) {
            logger.error("Колонка 'ППП (from file2)' не найдена.");
            return;
        }

        // Группировка строк по "ППП (from file2)"
        Map<String, List<Row>> groupedRows = new LinkedHashMap<>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(positionColumnIndex);
                String key = cell.getStringCellValue();
                groupedRows.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }
        }
        logger.info("Строки успешно сгруппированы по колонке 'ППП (from file2)'.");

        // Создание нового листа для сгруппированных данных
        Sheet outputSheet = sheet.getWorkbook().createSheet("GroupedData");

        // Запись строки заголовка
        Row newHeaderRow = outputSheet.createRow(0);
        for (int colIndex = 0; colIndex < headerRow.getLastCellNum(); colIndex++) {
            Cell oldCell = headerRow.getCell(colIndex);
            Cell newCell = newHeaderRow.createCell(colIndex);
            newCell.setCellValue(oldCell.getStringCellValue());
        }
        newHeaderRow.createCell(headerRow.getLastCellNum()).setCellValue("Кол-во");

        int currentRowNum = 1; // Начало со второй строки, так как первая строка - заголовок
        for (Map.Entry<String, List<Row>> entry : groupedRows.entrySet()) {
            String key = entry.getKey();
            List<Row> rows = entry.getValue();

            // Добавление заголовка группы
            Row groupHeaderRow = outputSheet.createRow(currentRowNum++);
            Cell groupHeaderCell = groupHeaderRow.createCell(0);
            groupHeaderCell.setCellValue("Код позиции (from file1): " + key);

            // Добавление строк
            for (Row row : rows) {
                Row newRow = outputSheet.createRow(currentRowNum++);
                for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                    Cell oldCell = row.getCell(colIndex);
                    Cell newCell = newRow.createCell(colIndex);
                    if (oldCell != null) {
                        switch (oldCell.getCellType()) {
                            case STRING:
                                newCell.setCellValue(oldCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                newCell.setCellValue(oldCell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                newCell.setCellValue(oldCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                newCell.setCellFormula(oldCell.getCellFormula());
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            // Добавление строки с количеством
            Row countRow = outputSheet.createRow(currentRowNum++);
            Cell countLabelCell = countRow.createCell(0);
            countLabelCell.setCellValue("Кол-во");
            Cell countValueCell = countRow.createCell(1);
            countValueCell.setCellValue(rows.size());
            logger.info("Группа '{}' содержит {} строк.", key, rows.size());
        }

        // Добавление строки с общим количеством записей в конце таблицы
        Row totalCountRow = outputSheet.createRow(currentRowNum++);
        Cell totalCountLabelCell = totalCountRow.createCell(0);
        totalCountLabelCell.setCellValue("Общее количество записей");
        Cell totalCountValueCell = totalCountRow.createCell(1);
        totalCountValueCell.setCellValue(sheet.getLastRowNum());
        logger.info("Общее количество записей: {}", sheet.getLastRowNum());
    }
}
