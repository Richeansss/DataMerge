package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class GroupRowsByPositionAndCount {

    public static void main(String[] args) throws IOException {
        String inputFilePath = "MergedData.xlsx";
        String outputFilePath = "SortedData_Output.xlsx";

        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("SortedData");

        int positionColumnIndex = -1;
        Row headerRow = sheet.getRow(0);

        // Find the "Код позиции (from file1)" column index
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equals("ППП (from file2)")) {
                positionColumnIndex = cell.getColumnIndex();
                break;
            }
        }

        if (positionColumnIndex == -1) {
            System.out.println("Column 'Код позиции (from file1)' not found.");
            workbook.close();
            fis.close();
            return;
        }

        // Group rows by "Код позиции (from file1)"
        Map<String, List<Row>> groupedRows = new LinkedHashMap<>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(positionColumnIndex);
                String key = cell.getStringCellValue();
                groupedRows.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }
        }


        // Create a new sheet for the grouped data
        Sheet outputSheet = workbook.createSheet("GroupedData");

        // Write the header row
        Row newHeaderRow = outputSheet.createRow(0);
        for (int colIndex = 0; colIndex < headerRow.getLastCellNum(); colIndex++) {
            Cell oldCell = headerRow.getCell(colIndex);
            Cell newCell = newHeaderRow.createCell(colIndex);
            newCell.setCellValue(oldCell.getStringCellValue());
        }
        newHeaderRow.createCell(headerRow.getLastCellNum()).setCellValue("Кол-во");

        int currentRowNum = 1; // Start from the second row because the first row is the header
        for (Map.Entry<String, List<Row>> entry : groupedRows.entrySet()) {
            String key = entry.getKey();
            List<Row> rows = entry.getValue();

            // Add header for the group
            Row groupHeaderRow = outputSheet.createRow(currentRowNum++);
            Cell groupHeaderCell = groupHeaderRow.createCell(0);
            groupHeaderCell.setCellValue("Код позиции (from file1): " + key);

            // Add rows
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

            // Add count row
            Row countRow = outputSheet.createRow(currentRowNum++);
            Cell countLabelCell = countRow.createCell(0);
            countLabelCell.setCellValue("Кол-во");
            Cell countValueCell = countRow.createCell(1);
            countValueCell.setCellValue(rows.size());
        }

        // Add total count row at the end of the table
        Row totalCountRow = outputSheet.createRow(currentRowNum++);
        Cell totalCountLabelCell = totalCountRow.createCell(0);
        totalCountLabelCell.setCellValue("Общее количество записей");
        Cell totalCountValueCell = totalCountRow.createCell(1);
        totalCountValueCell.setCellValue(sheet.getLastRowNum());

        // Save the modified workbook
        FileOutputStream fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);
        fos.close();
        workbook.close();
        fis.close();
    }
}
