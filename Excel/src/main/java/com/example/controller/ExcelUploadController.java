package com.example.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;

@RestController
@RequestMapping("/api/excel")
@CrossOrigin(origins = "http://localhost:4200") // Replace with your
public class ExcelUploadController {

    @PostMapping("/upload")
    public ResponseEntity<Object> uploadFiles(@RequestParam("file1") MultipartFile file1,
                                               @RequestParam("file2") MultipartFile file2) {
        List<Map<String, String>> file1Data = new ArrayList<>();
        List<Map<String, String>> file2Data = new ArrayList<>();
        Set<String> headers = new LinkedHashSet<>();
        Map<String, Map<String, String>> mergedData = new LinkedHashMap<>();
        Set<String> commonIds = new HashSet<>();

        try {
            // Get headers and data from both files
            headers.addAll(getHeadersAndData(file1, file1Data));
            headers.addAll(getHeadersAndData(file2, file2Data));
        } catch (IOException e) {
            return ResponseEntity.badRequest().body("Error processing files: " + e.getMessage());
        }

        // Identify common IDs between both files
        Set<String> file1Ids = extractIds(file1Data);
        Set<String> file2Ids = extractIds(file2Data);
        commonIds.addAll(file1Ids);
        commonIds.retainAll(file2Ids); // Keep only IDs that are present in both files

        // Merge data for common IDs
        mergeData(file1Data, mergedData, commonIds);
        mergeData(file2Data, mergedData, commonIds);

        return createMergedDataFile(mergedData, headers);
    }

    private void mergeData(List<Map<String, String>> fileData, Map<String, Map<String, String>> mergedData, Set<String> commonIds) {
        for (Map<String, String> rowData : fileData) {
            String id = rowData.get("ID");
            if (id != null && commonIds.contains(id)) { // Only merge rows with matching IDs
                mergedData.putIfAbsent(id, new HashMap<>());
                mergedData.get(id).putAll(rowData);
            }
        }
    }

    private Set<String> getHeadersAndData(MultipartFile file, List<Map<String, String>> dataList) throws IOException {
        Set<String> headers = new LinkedHashSet<>();
        Workbook workbook = new XSSFWorkbook(file.getInputStream());

        Sheet sheet = workbook.getSheetAt(0);
        Row headerRow = sheet.getRow(0);  // Assuming the first row contains headers

        if (headerRow != null) {
            // Loop through each cell in the header row and add only non-empty headers
            for (int j = 0; j < headerRow.getPhysicalNumberOfCells(); j++) {
                String header = getCellValue(headerRow.getCell(j)).trim();
                if (!header.isEmpty()) {  // Ignore blank columns
                    headers.add(header);
                }
            }

            // Loop through data rows starting from the second row (index 1)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Map<String, String> rowData = new HashMap<>();
                    int headerIndex = 0;  // Tracks the corresponding header

                    for (int j = 0; j < headerRow.getPhysicalNumberOfCells(); j++) {
                        String header = getCellValue(headerRow.getCell(j)).trim();
                        if (!header.isEmpty()) {  // Skip blank columns
                            String value = getCellValue(row.getCell(j));
                            rowData.put(header, value);
                        }
                    }
                    dataList.add(rowData);
                }
            }
        }

        workbook.close();
        return headers;
    }

    private Set<String> extractIds(List<Map<String, String>> fileData) {
        Set<String> ids = new HashSet<>();
        for (Map<String, String> rowData : fileData) {
            String id = rowData.get("ID");
            if (id != null && !id.isEmpty()) {
                ids.add(id);
            }
        }
        return ids;
    }

    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return String.valueOf(cell.getDateCellValue());
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    private ResponseEntity<Object> createMergedDataFile(Map<String, Map<String, String>> mergedData, Set<String> headers) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Merged Data");
        Row headerRow = sheet.createRow(0);
        int columnIndex = 0;

        // Create the header row
        for (String header : headers) {
            headerRow.createCell(columnIndex++).setCellValue(header);
        }

        // Populate the data rows
        int rowIndex = 1;
        for (Map<String, String> rowData : mergedData.values()) {
            Row row = sheet.createRow(rowIndex++);
            columnIndex = 0;
            for (String header : headers) {
                row.createCell(columnIndex++).setCellValue(rowData.getOrDefault(header, ""));
            }
        }

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            workbook.close();

            HttpHeaders httpHeaders = new HttpHeaders();
            httpHeaders.add("Content-Disposition", "attachment; filename=merged_data.xlsx");
            return ResponseEntity.ok()
                    .headers(httpHeaders)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(outputStream.toByteArray());
        } catch (IOException e) {
            return ResponseEntity.badRequest().body("Error creating the Excel file: " + e.getMessage());
        }
    }
}
