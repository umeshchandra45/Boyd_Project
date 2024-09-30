package com.Boyd.O2C.prac;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.Scanner;

public class ExcelMatcher3 {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // Read file paths from console
        System.out.print("Enter the path for the business file: ");
        String businessFile = scanner.nextLine();

        System.out.print("Enter the path for the query file: ");
        String queryFile = scanner.nextLine();

        System.out.print("Enter the name for the output file (without extension): ");
        String outputFileName = scanner.nextLine();

        String outputFile = "D:\\Work\\Output_Sheets\\" + outputFileName + "_output.xlsx";
        String primaryColumn = "Id";

        try {
            mainProcess(businessFile, queryFile, primaryColumn, outputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void mainProcess(String businessFile, String queryFile, String primaryColumn, String outputFile) throws IOException {
        Workbook businessWorkbook = new XSSFWorkbook(new FileInputStream(businessFile));
        Workbook queryWorkbook = new XSSFWorkbook(new FileInputStream(queryFile));

        Sheet businessSheet = businessWorkbook.getSheetAt(0);
        Sheet querySheet = queryWorkbook.getSheetAt(0);

        Map<String, Map<String, String>> businessData = readSheet(businessSheet, primaryColumn);
        Map<String, Map<String, String>> queryData = readSheet(querySheet, primaryColumn);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Matched Data");

        Set<String> baseColumns = new HashSet<>();
        businessData.values().forEach(row -> baseColumns.addAll(row.keySet()));
        queryData.values().forEach(row -> baseColumns.addAll(row.keySet()));

        createHeaderRow(outputSheet, baseColumns);

        int rowIndex = 1;
        int matchedCount = 0;
        int totalCount = 0;

        for (String key : businessData.keySet()) {
            Row row = outputSheet.createRow(rowIndex++);
            Map<String, String> businessRow = businessData.getOrDefault(key, new HashMap<>());
            Map<String, String> queryRow = queryData.getOrDefault(key, new HashMap<>());

            int cellIndex = 0;
            boolean matched = true;
            for (String column : baseColumns) {
                String businessValue = businessRow.getOrDefault(column, "").toUpperCase().trim();
                String queryValue = queryRow.getOrDefault(column, "").toUpperCase().trim();

                Cell cell = row.createCell(cellIndex++);
                cell.setCellValue(businessValue);

                cell = row.createCell(cellIndex++);
                cell.setCellValue(queryValue);

                cell = row.createCell(cellIndex++);
                if (businessValue.equals(queryValue)) {
                    cell.setCellValue("Matched");
                } else {
                    cell.setCellValue("Not Matched");
                    matched = false;
                }

                totalCount++;
                if (!businessValue.equals(queryValue)) {
                    matchedCount++;
                }
            }

            Cell resultCell = row.createCell(cellIndex);
            resultCell.setCellValue(matched ? "Pass" : "Fail");
        }

        // Add summary row for counts and error percentage
        Row summaryRow = outputSheet.createRow(rowIndex);
        Cell summaryCell = summaryRow.createCell(0);
        summaryCell.setCellValue("Total Matched: " + (totalCount - matchedCount));

        summaryCell = summaryRow.createCell(1);
        summaryCell.setCellValue("Total Not Matched: " + matchedCount);

        double errorPercent = (double) matchedCount * 100 / totalCount;
        summaryCell = summaryRow.createCell(2);
        summaryCell.setCellValue(String.format("Percent error: %.2f%%", errorPercent));

        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            outputWorkbook.write(fileOut);
        }

        businessWorkbook.close();
        queryWorkbook.close();
        outputWorkbook.close();

        System.out.println("File Created Successfully.");
    }

    public static Map<String, Map<String, String>> readSheet(Sheet sheet, String primaryColumn) {
        Map<String, Map<String, String>> data = new HashMap<>();
        int primaryColIndex = -1;

        Row headerRow = sheet.getRow(0);
        Map<Integer, String> headers = new HashMap<>();

        for (Cell cell : headerRow) {
            String header = cell.getStringCellValue();
            headers.put(cell.getColumnIndex(), header);
            if (header.equals(primaryColumn)) {
                primaryColIndex = cell.getColumnIndex();
            }
        }

        if (primaryColIndex == -1) {
            throw new IllegalArgumentException("Primary column not found in the sheet: " + primaryColumn);
        }

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue; // Skip empty rows

            String primaryValue = row.getCell(primaryColIndex).getStringCellValue();
            Map<String, String> rowData = new HashMap<>();

            for (Map.Entry<Integer, String> entry : headers.entrySet()) {
                Cell cell = row.getCell(entry.getKey());
                String cellValue = "";
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            cellValue = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            cellValue = String.valueOf(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        default:
                            cellValue = "";
                    }
                }
                rowData.put(entry.getValue(), cellValue);
            }

            data.put(primaryValue, rowData);
        }

        return data;
    }

    public static void createHeaderRow(Sheet sheet, Set<String> baseColumns) {
        Row headerRow = sheet.createRow(0);
        int cellIndex = 0;

        for (String column : baseColumns) {
            Cell cell = headerRow.createCell(cellIndex++);
            cell.setCellValue(column + "_Input");

            cell = headerRow.createCell(cellIndex++);
            cell.setCellValue(column + "_Output");

            cell = headerRow.createCell(cellIndex++);
            cell.setCellValue(column + "_Status");
        }

        // Adding Result column header
        headerRow.createCell(cellIndex).setCellValue("Result");
    }
}