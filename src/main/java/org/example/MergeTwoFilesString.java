package org.example;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class MergeTwoFilesString {
    private static final Logger logger = LoggerFactory.getLogger(ExcelMerger.class);

    public static void main(String[] args) {
        String file1Path = "ДТОиР студентам/Файлы для объединения/М.1.С,_комментарии_КР_04.07.2024.xlsx";
        String file2Path = "ДТОиР студентам/Файлы для объединения/М.1.С,_комментарии_ТОиТР_04.07.2024.xlsx";
        String outputFilePath = "ДТОиР студентам/Файлы для объединения/file.xlsx";

        // Устанавливаем минимальное отношение раздувания до 0.001 (по умолчанию 0.01)
        ZipSecureFile.setMinInflateRatio(0.001);

        try {
            mergeExcelFiles(file1Path, file2Path, outputFilePath, 19); // Пропускаем первую строку
        } catch (IOException e) {
            logger.error("Error merging Excel files", e);
        }
    }

    public static void mergeExcelFiles(String file1Path, String file2Path, String outputFilePath, int skipRows) throws IOException {
        try (FileInputStream fis1 = new FileInputStream(file1Path);
             FileInputStream fis2 = new FileInputStream(file2Path);
             Workbook workbook1 = new XSSFWorkbook(fis1);
             Workbook workbook2 = new XSSFWorkbook(fis2);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet sheet1 = workbook1.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(0);
            Sheet outputSheet = outputWorkbook.createSheet("Merged");

            logger.info("Начинаем копирование данных из первого файла...");
            int rowCount = copySheet(sheet1, outputSheet, 0, skipRows);

            logger.info("Начинаем копирование данных из второго файла...");
            copySheet(sheet2, outputSheet, rowCount, skipRows);

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
                logger.info("Файлы успешно объединены в {}", outputFilePath);
            }
        }
    }

    private static int copySheet(Sheet sourceSheet, Sheet destinationSheet, int startRow, int skipRows) {
        int rowCount = startRow;
        for (int i = sourceSheet.getFirstRowNum() + skipRows; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                Row newRow = destinationSheet.createRow(rowCount++);
                logger.info("Копируем строку {} из исходного файла в строку {} результирующего файла", sourceRow.getRowNum(), newRow.getRowNum());
                copyRow(sourceRow, newRow);
            }
        }
        return rowCount;
    }

    private static void copyRow(Row sourceRow, Row destinationRow) {
        for (Cell sourceCell : sourceRow) {
            Cell newCell = destinationRow.createCell(sourceCell.getColumnIndex());
            logger.info("Копируем ячейку ({}, {}) со значением {} в результирующий файл", sourceRow.getRowNum(), sourceCell.getColumnIndex(), getCellValue(sourceCell));
            copyCell(sourceCell, newCell);
        }
    }

    private static void copyCell(Cell sourceCell, Cell destinationCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                destinationCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                destinationCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                destinationCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                destinationCell.setBlank();
                break;
            default:
                break;
        }
        CellStyle newCellStyle = destinationCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        destinationCell.setCellStyle(newCellStyle);
    }

    private static String getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case BLANK -> "";
            default -> "";
        };
    }
}
