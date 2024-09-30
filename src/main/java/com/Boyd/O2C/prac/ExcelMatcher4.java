package com.Boyd.O2C.prac;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class ExcelMatcher4 {
    public static void main(String[] args) {
        String businessFile = "D:\\Work\\Input_Sheets\\Source-14246.xlsx";
        String queryFile = "D:\\Work\\Input_Sheets\\Target-14239.xlsx";
        String primaryColumn = "Id";
        String outputFile = "D:\\Work\\Output_Sheets\\OutputFile_output9.xlsx";

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

        Map<String, Map<String, String>> businessData = readSheet(businessSheet, primaryColumn, primaryColumn);
        Map<String, Map<String, String>> queryData = readSheet(querySheet, primaryColumn, primaryColumn);

        Workbook outputWorkbook = new SXSSFWorkbook(); // Use SXSSFWorkbook for streaming
        Sheet outputSheet = outputWorkbook.createSheet("Matched Data");

        Set<String> allKeys = new HashSet<>();
        allKeys.addAll(businessData.keySet());
        allKeys.addAll(queryData.keySet());

        Set<String> baseColumns = businessData.values().iterator().next().keySet();

        createHeaderRow(outputSheet, baseColumns);

        int rowIndex = 1;
        int totalRows = 0;
        int totalCells = 0;
        Map<String, Integer> statusCounts = new HashMap<>();

        for (String key : allKeys) {
            Row row = outputSheet.createRow(rowIndex++);
            Map<String, String> businessRow = businessData.getOrDefault(key, new HashMap<>());
            Map<String, String> queryRow = queryData.getOrDefault(key, new HashMap<>());

            int cellIndex = 0;
            boolean isFail = false;

            for (String column : baseColumns) {
                String businessValue = businessRow.getOrDefault(column, "").toUpperCase().trim();
                String queryValue = queryRow.getOrDefault(column, "").toUpperCase().trim();

                Cell cell = row.createCell(cellIndex++);
                cell.setCellValue(businessValue);

                cell = row.createCell(cellIndex++);
                cell.setCellValue(queryValue);

                cell = row.createCell(cellIndex++);
                String status = businessValue.equals(queryValue) ? "Matched" : "Not Matched";
                cell.setCellValue(status);

                if (!status.equals("Matched")) {
                    isFail = true;
                }

                totalCells++;
                // Count occurrences of each status
                statusCounts.put(status, statusCounts.getOrDefault(status, 0) + 1);
            }

            Cell resultCell = row.createCell(cellIndex);
            resultCell.setCellValue(isFail ? "Fail" : "Pass");
        }

        // Add summary rows for each status column
        Row summaryRow = outputSheet.createRow(rowIndex++);
        int colIndex = 0;
        for (String column : baseColumns) {
            colIndex += 2; // Skip the input and output columns
            int matchedCount = statusCounts.getOrDefault("Matched", 0);
            int errorCount = statusCounts.getOrDefault("Not Matched", 0);
            double errorPercentage = (totalCells > 0) ? (errorCount * 100.0) / totalCells : 0.0;

            // Ensure the summary is in the correct cell
            String summaryText = String.format("Total Matched: %d, Error: %d, Percent Error: %.2f%%", matchedCount, errorCount, errorPercentage);
            summaryRow.createCell(colIndex).setCellValue(summaryText);
        }

        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            outputWorkbook.write(fileOut);
        }

        businessWorkbook.close();
        queryWorkbook.close();
        outputWorkbook.close();

        System.out.println("File Created Successfully.............");
        System.out.println("\nNo of cells that did not match:- " + statusCounts.getOrDefault("Not Matched", 0));
        System.out.println("Percent error in the sheets:- " + ((statusCounts.getOrDefault("Not Matched", 0) * 100.0) / totalCells) + " %");
    }

    public static Map<String, Map<String, String>> readSheet(Sheet sheet, String originalPrimaryColumn, String newPrimaryColumn) {
        Map<String, Map<String, String>> data = new HashMap<>();
        int primaryColIndex = -1;

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            throw new IllegalArgumentException("Header row is missing in the sheet.");
        }

        Map<Integer, String> headers = new HashMap<>();

        for (Cell cell : headerRow) {
            String header = cell.getStringCellValue();
            headers.put(cell.getColumnIndex(), header);
            if (header.equals(originalPrimaryColumn)) {
                primaryColIndex = cell.getColumnIndex();
            }
        }

        if (primaryColIndex == -1) {
            throw new IllegalArgumentException("Primary column not found in the sheet: " + originalPrimaryColumn);
        }

        // Rename the primary column in headers
        headers.put(primaryColIndex, newPrimaryColumn);

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue; // Skip if the row is null
            Cell primaryCell = row.getCell(primaryColIndex);
            String primaryValue = primaryCell == null ? "" : primaryCell.toString();
            Map<String, String> rowData = new HashMap<>();

            for (Map.Entry<Integer, String> entry : headers.entrySet()) {
                Cell cell = row.getCell(entry.getKey());
                rowData.put(entry.getValue(), cell == null ? "" : cell.toString());
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

        // Add Result column
        headerRow.createCell(cellIndex).setCellValue("Result");
    }
}