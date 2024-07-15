package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class CountUniqueValues {

    public static Map<String, Integer> countUniqueValuesInColumn(String[] uniqueValues, String filePath, String sheetName) throws IOException {
        Set<String> uniqueValuesSet = prepareUniqueValuesSet(uniqueValues);

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);

        int columnIdx = findColumnIndex(sheet, "№ строки п/п");

        if (columnIdx == -1) {
            System.out.println("Column containing '№ строки п/п' not found.");
            workbook.close();
            fis.close();
            return Collections.emptyMap();
        }

        Map<String, Integer> occurrencesMap = new HashMap<>();
        // Count the unique values in the specified column
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(columnIdx);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String cellValue = removeLeadingApostrophe(cell.getStringCellValue());
                    System.out.println("Read cell value: " + cellValue);  // Отладочное сообщение
                    if (uniqueValuesSet.contains(cellValue)) {
                        occurrencesMap.put(cellValue, occurrencesMap.getOrDefault(cellValue, 0) + 1);
                    }
                }
            }
        }

        workbook.close();
        fis.close();

        // Print occurrences
        System.out.println("Unique Values Found in the Excel Sheet:");
        for (Map.Entry<String, Integer> entry : occurrencesMap.entrySet()) {
            System.out.println(entry.getKey() + ": " + entry.getValue());
        }

        return occurrencesMap;
    }

    private static Set<String> prepareUniqueValuesSet(String[] uniqueValues) {
        Set<String> uniqueValuesSet = new HashSet<>();
        for (String value : uniqueValues) {
            uniqueValuesSet.add(removeLeadingApostrophe(value));
        }
        return uniqueValuesSet;
    }

    private static String removeLeadingApostrophe(String value) {
        return value.startsWith("'") ? value.substring(1) : value;
    }

    private static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            System.out.println("Header row is null.");
            return -1;
        }

        for (Cell cell : headerRow) {
            if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().contains(columnName)) {
                return cell.getColumnIndex();
            }
        }
        System.out.println("Column with name '" + columnName + "' not found.");
        return -1;
    }
}