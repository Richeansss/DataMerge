package org.example;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * Утилитарный класс для работы с Excel файлами.
 */
public class ExcelUtils {
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * Извлекает данные из первой страницы рабочей книги и создает карту с ключами из указанной колонки.
     *
     * @param workbook      Рабочая книга Excel.
     * @param keyColumnIndex Индекс колонки, содержащей ключевые значения.
     * @return Карта, где ключ - значение ячейки из указанной колонки, значение - строка.
     */
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
        logger.info("Данные успешно извлечены из рабочей книги.");
        return dataMap;
    }

    /**
     * Создает строку заголовка в указанном листе на основе заголовков из двух рабочих книг.
     *
     * @param sheet    Лист, в который добавляется строка заголовка.
     * @param workbook1 Первая рабочая книга.
     * @param workbook2 Вторая рабочая книга.
     */
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
        logger.info("Строка заголовка успешно создана.");
    }

    /**
     * Копирует данные из одной строки в другую, начиная с указанного индекса.
     *
     * @param sourceRow Исходная строка.
     * @param targetRow Целевая строка.
     * @param startIndex Индекс, с которого начинается копирование.
     * @param sourceFile Источник файла.
     */
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

    /**
     * Устанавливает значение ячейки в целевую ячейку на основе значения исходной ячейки.
     *
     * @param sourceCell Исходная ячейка.
     * @param targetCell Целевая ячейка.
     */
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
                    targetCell.setCellValue(""); // Обработка пустых ячеек
                    break;
                default:
                    targetCell.setCellValue(sourceCell.toString());
            }
        } else {
            targetCell.setCellValue(""); // Обработка null ячеек
        }
    }

    /**
     * Возвращает значение ячейки в виде строки.
     *
     * @param cell Ячейка для получения значения.
     * @return Значение ячейки в виде строки.
     */
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

    /**
     * Применяет стили к колонкам листа, начиная с указанного индекса.
     *
     * @param sheet Лист для применения стилей.
     * @param startIndex Начальный индекс колонки.
     * @param numColumns Количество колонок для стилизации.
     * @param source Источник файла для различной стилизации.
     */
    public static void applyColumnStyles(Sheet sheet, int startIndex, int numColumns, String source) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        int actualNumColumns = Math.min(numColumns, sheet.getRow(0).getLastCellNum() - startIndex);

        for (int i = startIndex; i < startIndex + actualNumColumns; i++) {
            CellStyle columnStyle = workbook.createCellStyle();
            columnStyle.cloneStyleFrom(style);

            // Установка разных цветов на основе источника
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

    /**
     * Добавляет строку в лист с несоответствующими данными.
     *
     * @param workbook Рабочая книга Excel.
     * @param sourceRow Исходная строка.
     * @param sheetName Имя листа для добавления строки.
     */
    public static void addUnmatchedRow(Workbook workbook, Row sourceRow, String sheetName) {
        Sheet unmatchedSheet = workbook.getSheet(sheetName);
        if (unmatchedSheet == null) {
            unmatchedSheet = workbook.createSheet(sheetName);

            // Создание строки заголовка для листа несоответствующих данных на основе заголовков исходной строки
            Row headerRow = unmatchedSheet.createRow(0);
            for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                Cell sourceCell = sourceRow.getCell(i);
                Cell newCell = headerRow.createCell(i);
                newCell.setCellValue(getCellValueAsString(sourceCell));
            }
            logger.info("Создан новый лист с несоответствующими данными: {}", sheetName);
        } else {
            // Проверка, существует ли строка заголовка, создание при отсутствии
            if (unmatchedSheet.getRow(0) == null) {
                Row headerRow = unmatchedSheet.createRow(0);
                for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                    Cell sourceCell = sourceRow.getCell(i);
                    Cell newCell = headerRow.createCell(i);
                    newCell.setCellValue(getCellValueAsString(sourceCell));
                }
                logger.info("Создана строка заголовка для листа: {}", sheetName);
            }
        }

        int rowIndex = unmatchedSheet.getLastRowNum() + 1;
        Row newRow = unmatchedSheet.createRow(rowIndex);
        copyRowData(sourceRow, newRow, 0, ""); // Путь к исходному файлу не нужен для несоответствующих строк
    }
}
