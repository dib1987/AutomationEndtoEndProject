package com.qa.opencart.utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Utility class for reading test data from Excel (.xlsx) files.
 * Supports multi-sheet access, type-safe cell reading, row count queries,
 * and map-based data access for cleaner TestNG DataProvider integration.
 */
public class ExcelUtil {

    private static final Logger log = LoggerFactory.getLogger(ExcelUtil.class);

    private static final String TEST_DATA_SHEET_PATH =
            "./src/test/resources/testdata/opencarttestdata.xlsx";

    // Private constructor — utility class, no instantiation needed
    private ExcelUtil() {}

    // ─────────────────────────────────────────────────────────────
    // Core: 2D Object array (TestNG DataProvider compatible)
    // ─────────────────────────────────────────────────────────────

    /**
     * Returns a 2D Object array of test data from the specified sheet,
     * skipping the header row. Suitable for use with @DataProvider.
     *
     * @param sheetName name of the sheet to read
     * @return 2D Object[][] with all data rows
     */
    public static Object[][] getTestData(String sheetName) {
        return getTestData(TEST_DATA_SHEET_PATH, sheetName);
    }

    /**
     * Overloaded version allowing a custom file path.
     *
     * @param filePath  path to the Excel file
     * @param sheetName name of the sheet to read
     * @return 2D Object[][] with all data rows
     */
    public static Object[][] getTestData(String filePath, String sheetName) {
        log.info("Reading test data from sheet '{}' in file '{}'", sheetName, filePath);

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = getSheet(workbook, sheetName);
            int rowCount = sheet.getLastRowNum();           // excludes header
            int colCount = sheet.getRow(0).getLastCellNum();

            Object[][] data = new Object[rowCount][colCount];

            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.getRow(i + 1);
                if (row == null) continue;
                for (int j = 0; j < colCount; j++) {
                    data[i][j] = getCellValueAsString(row.getCell(j));
                }
            }

            log.info("Read {} rows × {} columns from sheet '{}'", rowCount, colCount, sheetName);
            return data;

        } catch (IOException e) {
            log.error("Failed to read Excel file '{}': {}", filePath, e.getMessage());
            throw new RuntimeException("ExcelUtil: Unable to read file: " + filePath, e);
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Map-based access: column header → cell value
    // ─────────────────────────────────────────────────────────────

    /**
     * Returns a List of Maps where each map represents one data row,
     * keyed by column headers from row 0.
     * Useful for accessing values by column name rather than index.
     *
     * @param sheetName name of the sheet to read
     * @return List of row maps: [{Username=admin, Password=1234}, ...]
     */
    public static List<Map<String, String>> getTestDataAsMaps(String sheetName) {
        return getTestDataAsMaps(TEST_DATA_SHEET_PATH, sheetName);
    }

    public static List<Map<String, String>> getTestDataAsMaps(String filePath, String sheetName) {
        log.info("Reading map-based data from sheet '{}'", sheetName);

        List<Map<String, String>> result = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = getSheet(workbook, sheetName);
            Row headerRow = sheet.getRow(0);
            int colCount = headerRow.getLastCellNum();

            // Build header index
            String[] headers = new String[colCount];
            for (int j = 0; j < colCount; j++) {
                headers[j] = getCellValueAsString(headerRow.getCell(j));
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> rowMap = new LinkedHashMap<>();
                for (int j = 0; j < colCount; j++) {
                    rowMap.put(headers[j], getCellValueAsString(row.getCell(j)));
                }
                result.add(rowMap);
            }

        } catch (IOException e) {
            log.error("Failed to read sheet '{}': {}", sheetName, e.getMessage());
            throw new RuntimeException("ExcelUtil: Unable to read sheet: " + sheetName, e);
        }

        return result;
    }

    // ─────────────────────────────────────────────────────────────
    // Utility: row/column count
    // ─────────────────────────────────────────────────────────────

    /**
     * Returns the number of data rows in the sheet (excludes header row).
     */
    public static int getRowCount(String sheetName) {
        return getRowCount(TEST_DATA_SHEET_PATH, sheetName);
    }

    public static int getRowCount(String filePath, String sheetName) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {
            return getSheet(workbook, sheetName).getLastRowNum();
        } catch (IOException e) {
            throw new RuntimeException("ExcelUtil: Unable to get row count for: " + sheetName, e);
        }
    }

    /**
     * Returns the number of columns based on the header row.
     */
    public static int getColumnCount(String sheetName) {
        return getColumnCount(TEST_DATA_SHEET_PATH, sheetName);
    }

    public static int getColumnCount(String filePath, String sheetName) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {
            return getSheet(workbook, sheetName).getRow(0).getLastCellNum();
        } catch (IOException e) {
            throw new RuntimeException("ExcelUtil: Unable to get column count for: " + sheetName, e);
        }
    }

    // ─────────────────────────────────────────────────────────────
    // Private helpers
    // ─────────────────────────────────────────────────────────────

    /**
     * Returns a Sheet by name, throwing a clear error if not found.
     */
    private static Sheet getSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException(
                    "ExcelUtil: Sheet '" + sheetName + "' not found. "
                    + "Available sheets: " + getAvailableSheets(workbook));
        }
        return sheet;
    }

    /**
     * Lists all sheet names in the workbook for error messages.
     */
    private static String getAvailableSheets(Workbook workbook) {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            names.add(workbook.getSheetName(i));
        }
        return names.toString();
    }

    /**
     * Safely reads any cell type and returns it as a trimmed String.
     * Handles: NUMERIC (including dates), STRING, BOOLEAN, FORMULA, BLANK, null.
     */
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        DataFormatter formatter = new DataFormatter();

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatter.formatCellValue(cell);  // preserves date format
                }
                // Avoid scientific notation for large numbers (e.g., phone numbers, IDs)
                double numVal = cell.getNumericCellValue();
                if (numVal == Math.floor(numVal) && !Double.isInfinite(numVal)) {
                    return String.valueOf((long) numVal);
                }
                return String.valueOf(numVal);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return formatter.formatCellValue(cell);
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}