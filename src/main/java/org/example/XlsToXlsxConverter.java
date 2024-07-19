package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class XlsToXlsxConverter {
    private static final Logger logger = LoggerFactory.getLogger(XlsToXlsxConverter.class);

    public static void convertXlsToXlsx(String inputFilePath, String outputFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook xlsWorkbook = new HSSFWorkbook(fis);
             Workbook xlsxWorkbook = new XSSFWorkbook()) {

            logger.info("Входной файл XLS успешно открыт.");

            for (int i = 0; i < xlsWorkbook.getNumberOfSheets(); i++) {
                Sheet xlsSheet = xlsWorkbook.getSheetAt(i);
                Sheet xlsxSheet = xlsxWorkbook.createSheet(xlsSheet.getSheetName());
                logger.info("Копирование листа: {}", xlsSheet.getSheetName());
                copySheet(xlsSheet, xlsxSheet);
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                xlsxWorkbook.write(fos);
                logger.info("Успешная запись в выходной файл XLSX.");
            }

            logger.info("Конвертация успешно завершена.");

        } catch (IOException e) {
            logger.error("Произошла ошибка при конвертации.", e);
            throw e;
        }
    }

    private static void copySheet(Sheet xlsSheet, Sheet xlsxSheet) {
        for (int i = 0; i < xlsSheet.getPhysicalNumberOfRows(); i++) {
            Row xlsRow = xlsSheet.getRow(i);
            Row xlsxRow = xlsxSheet.createRow(i);

            if (xlsRow != null) {
                for (int j = 0; j < xlsRow.getPhysicalNumberOfCells(); j++) {
                    Cell xlsCell = xlsRow.getCell(j);
                    Cell xlsxCell = xlsxRow.createCell(j);

                    if (xlsCell != null) {
                        copyCell(xlsCell, xlsxCell);
                    }
                }
            }
        }
    }

    private static void copyCell(Cell xlsCell, Cell xlsxCell) {
        switch (xlsCell.getCellType()) {
            case STRING:
                xlsxCell.setCellValue(xlsCell.getStringCellValue());
                break;
            case NUMERIC:
                xlsxCell.setCellValue(xlsCell.getNumericCellValue());
                break;
            case BOOLEAN:
                xlsxCell.setCellValue(xlsCell.getBooleanCellValue());
                break;
            case FORMULA:
                xlsxCell.setCellFormula(xlsCell.getCellFormula());
                break;
            case BLANK:
                xlsxCell.setBlank();
                break;
            default:
                break;
        }
    }
}