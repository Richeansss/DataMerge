package org.example;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class TableColumnSorter {
    private static final Logger logger = LoggerFactory.getLogger(TableColumnSorter.class);

    public static void sortColumnsByHeaders(Workbook workbook, String sheetName) {
        // Пары заголовков, которые нужно расположить рядом
        List<String[]> headerPairs = Arrays.asList(
                new String[]{"Структурное подразделение", "Подразделение ДО"},
                new String[]{"Код позиции", "Код подрядчика"},
                new String[]{"Инвентарный номер", "Инм.№"},
                new String[]{"Наименование МТР", "Полное наименование"},
                new String[]{"Содержание работ.Сводный код XYZ", "XYZ (from file2)"},
                new String[]{"Объект работ", "Наименование объекта"},
                new String[]{"Содержание работ (from file1)", "Содержание работ (from file2)"},
                new String[]{"Спецификация поставщика (from file1)", "№ спец (from file2)"},
                new String[]{"Поставщик (from file1)", "Поставщик (from file2)"},
                new String[]{"Доходный договор (from file1)", "№ договора подряда (from file2)"},
                new String[]{"Принципал", "ДО (from file2)"}
                // Добавьте больше пар по необходимости
        );

        // Получаем лист по имени
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Лист с именем " + sheetName + " не найден в рабочей книге.");
        }

        // Проверяем наличие листа "SortedData" или создаем новый
        Sheet sortedSheet = workbook.getSheet("SortedData");
        if (sortedSheet == null) {
            logger.info("Лист 'SortedData' не найден, создаем новый лист...");
            sortedSheet = workbook.createSheet("SortedData");
            logger.info("Создан новый лист 'SortedData'.");
        } else {
            logger.info("Найден существующий лист 'SortedData'.");
        }

        // Получаем заголовки из первой строки
        Row headerRow = sheet.getRow(0);
        Map<String, Integer> headerIndexMap = new HashMap<>();

        // Создаем карту соответствия заголовков и их индексов
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            String cellValue = getCellValueAsString(cell);
            headerIndexMap.put(cellValue, i);
        }

        // Список для хранения пар с индексами
        List<int[]> indexPairs = new ArrayList<>();

        // Перебираем пары заголовков и определяем индексы
        for (String[] pair : headerPairs) {
            String header1 = pair[0];
            String header2 = pair[1];

            int index1 = findHeaderIndex(headerIndexMap, header1);
            int index2 = findHeaderIndex(headerIndexMap, header2);

            if (index1 != -1 && index2 != -1) {
                indexPairs.add(new int[]{index1, index2});
            } else {
                logger.warn("Заголовки '{}' и '{}' не найдены в листе '{}'", header1, header2, sheetName);
            }
        }

        // Перемещаем пары столбцов в начало
        int targetIndex = 0;
        for (int[] indexes : indexPairs) {
            int index1 = indexes[0];
            int index2 = indexes[1];
            moveColumn(sheet, sortedSheet, index1, targetIndex++);
            moveColumn(sheet, sortedSheet, index2, targetIndex++);
        }

        // Перемещаем оставшиеся столбцы
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            if (!isIndexInPairs(i, indexPairs)) {
                moveColumn(sheet, sortedSheet, i, targetIndex++);
            }
        }

        // Записываем рабочую книгу в файл
        try (FileOutputStream fileOut = new FileOutputStream("MergedData.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            logger.error("Ошибка записи рабочей книги в файл", e);
        }
    }

    private static boolean isIndexInPairs(int index, List<int[]> indexPairs) {
        for (int[] pair : indexPairs) {
            if (pair[0] == index || pair[1] == index) {
                return true;
            }
        }
        return false;
    }

    private static void moveColumn(Sheet sourceSheet, Sheet targetSheet, int sourceIndex, int targetIndex) {
        int lastRow = sourceSheet.getLastRowNum();

        for (int i = 0; i <= lastRow; i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row targetRow = targetSheet.getRow(i);

            if (sourceRow == null) {
                continue; // Пропускаем пустые строки в исходном листе
            }

            if (targetRow == null) {
                targetRow = targetSheet.createRow(i); // Создаем строку в целевом листе, если она отсутствует
            }

            // Переносим столбец
            Cell sourceCell = sourceRow.getCell(sourceIndex);
            Cell targetCell = targetRow.createCell(targetIndex);

            if (sourceCell != null) {
                setCellValue(sourceCell, targetCell);
            } else {
                targetCell.setCellValue(""); // Обработка пустых ячеек
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
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> Double.toString(cell.getNumericCellValue());
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case BLANK -> "";
            default -> cell.toString();
        };
    }

    private static int findHeaderIndex(Map<String, Integer> headerIndexMap, String header) {
        for (Map.Entry<String, Integer> entry : headerIndexMap.entrySet()) {
            if (entry.getKey().contains(header)) {
                return entry.getValue();
            }
        }
        return -1;
    }

}
