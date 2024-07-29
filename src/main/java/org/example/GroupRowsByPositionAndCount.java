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
 * Класс для группировки строк в Excel-файле по указанному полю и подсчета количества строк в каждой группе.
 */
public class GroupRowsByPositionAndCount {
    private static final Logger logger = LoggerFactory.getLogger(GroupRowsByPositionAndCount.class);

    /**
     * Группирует строки и подсчитывает их количество в одном и том же файле.
     *
     * @param inputFilePath      Путь к входному файлу Excel.
     * @param sourceSheetName    Имя листа, на котором будет выполняться группировка.
     * @param targetSheetName    Имя листа, на который будут записаны результаты группировки.
     * @param groupingColumnName Имя столбца, по которому будет происходить группировка.
     * @throws IOException Если возникнет ошибка при чтении или записи файла.
     */
    public static void groupRowsAndCountInSameFile(String inputFilePath, String sourceSheetName, String targetSheetName, String groupingColumnName) throws IOException {
        logger.info("Чтение файла: {}", inputFilePath);
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sourceSheet = workbook.getSheet(sourceSheetName);

        if (sourceSheet == null) {
            logger.error("Лист '{}' не найден в файле '{}'.", sourceSheetName, inputFilePath);
            workbook.close();
            fis.close();
            return;
        }

        groupAndCountRows(sourceSheet, workbook, targetSheetName, groupingColumnName);

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
     * @param sourceSheet        Лист, на котором будет выполняться группировка.
     * @param workbook           Рабочая книга Excel.
     * @param targetSheetName    Имя листа, на который будут записаны результаты группировки.
     * @param groupingColumnName Имя столбца, по которому будет происходить группировка.
     */
    public static void groupAndCountRows(Sheet sourceSheet, Workbook workbook, String targetSheetName, String groupingColumnName) {
        int groupingColumnIndex = -1;
        Row headerRow = sourceSheet.getRow(0);

        // Поиск индекса колонки для группировки
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equals(groupingColumnName)) {
                groupingColumnIndex = cell.getColumnIndex();
                break;
            }
        }

        if (groupingColumnIndex == -1) {
            logger.error("Колонка '{}' не найдена.", groupingColumnName);
            return;
        }

        // Группировка строк по указанной колонке
        Map<String, List<Row>> groupedRows = new LinkedHashMap<>();
        for (int rowIndex = 1; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
            Row row = sourceSheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(groupingColumnIndex);
                String key = cell.getStringCellValue();
                groupedRows.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }
        }
        logger.info("Строки успешно сгруппированы по колонке '{}'.", groupingColumnName);

        // Создание нового листа для сгруппированных данных
        Sheet targetSheet = workbook.getSheet(targetSheetName);
        if (targetSheet == null) {
            targetSheet = workbook.createSheet(targetSheetName);
        } else {
            // Очистка существующего листа
            for (int i = targetSheet.getLastRowNum(); i >= 0; i--) {
                targetSheet.removeRow(targetSheet.getRow(i));
            }
        }

        // Запись строки заголовка
        Row newHeaderRow = targetSheet.createRow(0);
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
            Row groupHeaderRow = targetSheet.createRow(currentRowNum++);
            Cell groupHeaderCell = groupHeaderRow.createCell(0);
            groupHeaderCell.setCellValue("Код позиции: " + key);

            // Добавление строк
            for (Row row : rows) {
                Row newRow = targetSheet.createRow(currentRowNum++);
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
            Row countRow = targetSheet.createRow(currentRowNum++);
            Cell countLabelCell = countRow.createCell(0);
            countLabelCell.setCellValue("Кол-во");
            Cell countValueCell = countRow.createCell(1);
            countValueCell.setCellValue(rows.size());
            logger.info("Группа '{}' содержит {} строк.", key, rows.size());
        }

        // Добавление строки с общим количеством записей в конце таблицы
        Row totalCountRow = targetSheet.createRow(currentRowNum++);
        Cell totalCountLabelCell = totalCountRow.createCell(0);
        totalCountLabelCell.setCellValue("Общее количество записей");
        Cell totalCountValueCell = totalCountRow.createCell(1);
        totalCountValueCell.setCellValue(sourceSheet.getLastRowNum());
        logger.info("Общее количество записей: {}", sourceSheet.getLastRowNum());
    }
}
