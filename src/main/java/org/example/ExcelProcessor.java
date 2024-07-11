package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelProcessor {
    private static final Logger logger = LoggerFactory.getLogger(ExcelProcessor.class);

    public static void main(String[] args) {
        String filePath = "path/to/your/excel/file.xlsx";
        String sheetName = "Sheet1";
        int startRow = 0; // начальная строка
        int endRow = 100; // конечная строка
        int targetColumn = 5; // целевой столбец (6-й столбец в Excel)

        processExcelFile(filePath, sheetName, startRow, endRow, targetColumn);
    }

    public static void processExcelFile(String filePath, String sheetName, int startRow, int endRow, int targetColumn) {
        double sumLeftColumnGreaterEqual20000 = 0;
        double sumLeftColumnLess20000 = 0;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                logger.error("Sheet '{}' not found in the file '{}'", sheetName, filePath);
                return;
            }

            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                Cell targetCell = row.getCell(targetColumn);
                Cell leftCell = row.getCell(targetColumn - 2);
                if (targetCell == null || targetCell.getCellType() != CellType.NUMERIC) continue;
                if (leftCell == null || leftCell.getCellType() != CellType.NUMERIC) continue;

                double targetValue = targetCell.getNumericCellValue();
                double leftValue = leftCell.getNumericCellValue();

                if (targetValue >= 20000) {
                    sumLeftColumnGreaterEqual20000 += leftValue;
                    logger.info("Value {} found at row {}, column {} is >= 20000; adding left value {} to sum", targetValue, rowIndex + 1, targetColumn + 1, leftValue);
                } else {
                    sumLeftColumnLess20000 += leftValue;
                    logger.info("Value {} found at row {}, column {} is < 20000; adding left value {} to sum", targetValue, rowIndex + 1, targetColumn + 1, leftValue);
                }
            }

            logger.info("Sum of left column values where target column value >= 20000: {}", sumLeftColumnGreaterEqual20000);
            logger.info("Sum of left column values where target column value < 20000: {}", sumLeftColumnLess20000);

        } catch (IOException e) {
            logger.error("Ошибка записи в файл '{}'", filePath, e);
        }
    }
}
