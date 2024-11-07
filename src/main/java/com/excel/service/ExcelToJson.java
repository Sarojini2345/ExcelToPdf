package com.excel.service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.fasterxml.jackson.databind.ObjectMapper;

@Component
public class ExcelToJson {
	
	public static String convertExcelToJson(String excelFilePath) {
	    Map<String, Object> jsonOutput = new HashMap<>();

	    try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
	         Workbook workbook = createWorkbook(fis, excelFilePath)) {

	        List<Map<String, Object>> sheetsData = new ArrayList<>();

	        // Handle the first sheet separately
	        Map<String, Object> firstSheetData = processFirstSheet(workbook.getSheetAt(0));
	        sheetsData.add(firstSheetData);

	        // Handle the remaining sheets
	        for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
	            Map<String, Object> sheetData = processOtherSheets(workbook, sheetIndex);
	            sheetsData.add(sheetData);
	        }

	        // Add the sheets data to JSON output
	        jsonOutput.put("sheets", sheetsData);

	        // Convert to JSON
	        ObjectMapper objectMapper = new ObjectMapper();
	        return objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(jsonOutput);

	    } catch (IOException e) {
	        e.printStackTrace();
	        return null;
	    }
	}

	// Method to process the first sheet
	private static Map<String, Object> processFirstSheet(Sheet sheet) {
	    Map<String, Object> sheetData = new HashMap<>();
	    List<Map<String, String>> activities = new ArrayList<>();

	    // Add sheet name to JSON structure
	    sheetData.put("sheetName", sheet.getSheetName());

	    // Read project details from the first four rows, considering merged cells
	   // sheetData.put("projectName", getMergedCellValue(sheet, 0, 2).split(":").length > 1 ? getMergedCellValue(sheet, 0, 2).split(":")[1].trim() : "");
	   // sheetData.put("EmployeeName", getMergedCellValue(sheet, 1, 2).split(":").length > 1 ? getMergedCellValue(sheet, 1, 2).split(":")[1].trim() : "");
	  //  sheetData.put("PO Number", getMergedCellValue(sheet, 2, 2).split(":").length > 1 ? getMergedCellValue(sheet, 2, 2).split(":")[1].trim() : "");

	    // Iterate through the rows starting from row 6
	    for (int i = 2; i <=sheet.getLastRowNum(); i++) {
	        Row row = sheet.getRow(i);
	        if (row == null) continue; // Skip empty rows

	        Map<String, String> activity = new HashMap<>();
	        activity.put("Work Package", getCellValue(row.getCell(0)));
	        activity.put("Resource Name", getCellValue(row.getCell(1)));
	        activity.put("Role", getCellValue(row.getCell(2)));
	        activity.put("Location", getCellValue(row.getCell(3)));
	        activity.put("Daily Rate", getCellValue(row.getCell(4)));
	        activity.put("Hourly Rate", getCellValue(row.getCell(5)));
	        activity.put("Hours", getCellValue(row.getCell(6)));

	        String backgroundColor = getBackgroundColor(row.getCell(0), sheet.getWorkbook());
	        activity.put("backgroundColor", backgroundColor);

	        activities.add(activity);
	    }

	    // Add activities and total hours to sheet data
	    sheetData.put("timesheets", activities);
	  //  sheetData.put("totalhours", getCellValue(sheet.getRow(sheet.getLastRowNum()).getCell(3)));

	    return sheetData;
	}

	// Method to process other sheets
	private static Map<String, Object> processOtherSheets(Workbook workbook, int sheetIndex) {
	    Sheet sheet = workbook.getSheetAt(sheetIndex);
	    Map<String, Object> sheetData = new HashMap<>();
	    List<Map<String, String>> activities = new ArrayList<>();

	    // Add sheet name to JSON structure
	    sheetData.put("sheetName", sheet.getSheetName());

	    // Read project details from the first four rows, considering merged cells
	    sheetData.put("projectName", getMergedCellValue(sheet, 0, 2).split(":").length > 1 ? getMergedCellValue(sheet, 0, 2).split(":")[1].trim() : "");
	    sheetData.put("EmployeeName", getMergedCellValue(sheet, 1, 2).split(":").length > 1 ? getMergedCellValue(sheet, 1, 2).split(":")[1].trim() : "");
	    sheetData.put("PO Number", getMergedCellValue(sheet, 2, 2).split(":").length > 1 ? getMergedCellValue(sheet, 2, 2).split(":")[1].trim() : "");

	    // Iterate through the rows starting from row 6
	    for (int i = 5; i < sheet.getLastRowNum(); i++) {
	        Row row = sheet.getRow(i);
	        if (row == null) continue; // Skip empty rows

	        Map<String, String> activity = new HashMap<>();
	        activity.put("Slno", getCellValue(row.getCell(0)).replaceAll("\\.0$", ""));
	        activity.put("Date", getCellValue(row.getCell(1)));
	        activity.put("Hours", getCellValue(row.getCell(2)));
	        activity.put("work package Name", getCellValue(row.getCell(3)));
	        activity.put("Activities", getCellValue(row.getCell(4)));

	        String backgroundColor = getBackgroundColor(row.getCell(0), workbook);
	        activity.put("backgroundColor", backgroundColor);

	        activities.add(activity);
	    }

	    // Add activities and total hours to sheet data
	    sheetData.put("timesheets", activities);
	    sheetData.put("totalhours", getCellValue(sheet.getRow(sheet.getLastRowNum()).getCell(3)));

	    return sheetData;
	}

    // Existing helper methods
    private static Workbook createWorkbook(FileInputStream fis, String filePath) throws IOException {
        if (filePath.endsWith(".xlsx")) {
            return new XSSFWorkbook(fis);
        } else if (filePath.endsWith(".xls")) {
            return new HSSFWorkbook(fis);
        } else {
            throw new IllegalArgumentException("The specified file is not an Excel file.");
        }
    }

    private static String getMergedCellValue(Sheet sheet, int rowIndex, int columnIndex) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                Row row = sheet.getRow(mergedRegion.getFirstRow());
                Cell cell = row.getCell(mergedRegion.getFirstColumn());
                return getCellValue(cell);
            }
        }
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        return getCellValue(cell);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        String cellValue = "";
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
                break;
        }

        return cellValue.replaceAll("[\t\n]", "").trim();
    }

    private static String getBackgroundColor(Cell cell, Workbook workbook) {
        if (cell == null) return "";

        CellStyle style = cell.getCellStyle();
        if (style != null) {
            if (style instanceof XSSFCellStyle) {
                XSSFColor color = (XSSFColor) ((XSSFCellStyle) style).getFillForegroundColorColor();
                if (color != null) {
                    byte[] rgb = color.getRGB();
                    if (rgb != null) {
                        return String.format("#%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
                    }
                }
            } else if (style instanceof HSSFCellStyle) {
                HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
                short colorIndex = style.getFillForegroundColor();
                if (style.getFillPattern() != FillPatternType.NO_FILL) {
                    return String.format("#%02X%02X%02X",
                            palette.getColor(colorIndex).getTriplet()[0],
                            palette.getColor(colorIndex).getTriplet()[1],
                            palette.getColor(colorIndex).getTriplet()[2]);
                }
            }
        }
        return "";
    }


}
