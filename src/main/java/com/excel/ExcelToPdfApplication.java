package com.excel;

import java.io.IOException;
import java.net.MalformedURLException;

import com.excel.service.ExcelToJson;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;

import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.*;
import com.itextpdf.layout.property.HorizontalAlignment;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.UnitValue;
import com.itextpdf.layout.property.VerticalAlignment;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

public class ExcelToPdfApplication {

    public static void main(String[] args) {
    	
    	String excelFilePath = "C:\\Users\\hp\\Downloads\\Employees_Timesheets_Siemens Mobility GmbH_Service PLM SMO_01-10-2024-to-31-10-2024_06_Nov_2024_03_49_35.xls"; // Replace with your Excel file path
        String jsonString = ExcelToJson.convertExcelToJson(excelFilePath);
        System.out.println(jsonString);
        createPdfFromJsonSheets(jsonString);
    }
    
    public static void createPdfFromJsonSheets(String jsonString) {
        try {
            // Parse JSON data
            ObjectMapper objectMapper = new ObjectMapper();
            Map<String, Object> jsonData = objectMapper.readValue(jsonString, Map.class);

            // Extract all sheets
            List<Map<String, Object>> sheets = (List<Map<String, Object>>) jsonData.get("sheets");

            // Determine the file path for the Downloads directory
            String downloadsPath = System.getProperty("user.home") + File.separator + "Downloads";
            String pdfFilePath = downloadsPath + File.separator + "Timesheet.pdf";

            // Initialize PdfWriter and PdfDocument for the entire PDF
            PdfWriter writer = new PdfWriter(new FileOutputStream(pdfFilePath));
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document document = new Document(pdfDoc);

            // Set the document font size
            document.setFontSize(7);

            // Handle the first sheet separately with a unique design
            if (!sheets.isEmpty()) {
                Map<String, Object> firstSheet = sheets.get(0);
                addFirstSheetToPdf(document, firstSheet);
                // Add a page break after the first sheet
                document.add(new AreaBreak());
            }

            // Loop through the rest of the sheets and add them in a uniform design
            for (int i = 1; i < sheets.size(); i++) {
                Map<String, Object> sheet = sheets.get(i);
                addSheetToPdf(document, sheet);
                if (i < sheets.size() - 1) {
                    // Add a page break after each sheet except the last one
                    document.add(new AreaBreak());
                }
            }

            // Close the document
            document.close();
            System.out.println("PDF generated and saved to Downloads directory.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to add the first sheet with a unique design
    private static void addFirstSheetToPdf(Document document, Map<String, Object> sheetData) throws IOException {
        // Implement your custom design logic for the first sheet here
        // Adjusted column widths to fit within one page
    	 // Load the image (logo) - replace with the correct path
        String logoPath = "C:\\Users\\hp\\Downloads\\Intelizign Logo\\image 3.png";
        ImageData imageData = ImageDataFactory.create(logoPath);
        Image logo = new Image(imageData);
        logo.setWidth(100); // Set appropriate size for the logo
        logo.setHeight(20);

        // Create header table with two columns
        Table headerTable = new Table(UnitValue.createPercentArray(new float[]{70, 30}));
        headerTable.setWidth(UnitValue.createPercentValue(100));
        headerTable.setFixedLayout(); // Set layout for the header table
        headerTable.setBorder(new SolidBorder(1f));
        headerTable.setHeight(UnitValue.createPercentValue(100));

        // Add the left-aligned header text cell
        Cell leftCell = new Cell()
                .add(new Paragraph("Timesheets Summary October '24")
                        .setBold()
                        .setFontSize(12)
                        .setTextAlignment(TextAlignment.LEFT))
                .setBorder(null)
                .setPaddingRight(10)
                .setVerticalAlignment(VerticalAlignment.MIDDLE);
        headerTable.addCell(leftCell);

        // Add the right-aligned logo cell
        Cell rightCell = new Cell()
                .add(logo.setHorizontalAlignment(HorizontalAlignment.RIGHT))
                .setBorder(null)
                .setPaddingLeft(20)
                .setVerticalAlignment(VerticalAlignment.MIDDLE);
        headerTable.addCell(rightCell);

        // Add the header table to the document
        document.add(headerTable);
        Paragraph spacer = new Paragraph("").setMarginBottom(10); // Adjust the margin as needed
        document.add(spacer);
    	
        float[] columnWidths = {80, 80, 60, 60, 60, 60, 50};
        Table table = new Table(UnitValue.createPercentArray(columnWidths));
        table.setBorder(new SolidBorder(1f));
        // Define header color
        Color headerColor = new DeviceRgb(255, 255, 255); // Red color

        // Set a smaller font size for the entire document
        document.setFontSize(7);

        // Create header cells
        table.addCell(createHeaderCell("Work Package", headerColor));
        table.addCell(createHeaderCell("Resource Name", headerColor));

        // Merged header for "Finance"
        Paragraph financeMainHeader = new Paragraph("Finance")
                .setFontColor(ColorConstants.RED)
                .setTextAlignment(TextAlignment.CENTER)
                .setVerticalAlignment(VerticalAlignment.MIDDLE)
                .setFontSize(10); // Adjust font size for the main header



        // Create a new Table to hold the subheadings side by side
        Table subheadingTable = new Table(4); // 4 columns for 4 subheadings
        subheadingTable.setWidth(UnitValue.createPercentValue(100)); // Set width to 100%
        subheadingTable.setFixedLayout();

        // Add subheading cells in the same row
        subheadingTable.addCell(createSubheadingCell("Role"));
        subheadingTable.addCell(createSubheadingCell("Location"));
        subheadingTable.addCell(createSubheadingCell("Daily Rate"));
        subheadingTable.addCell(createSubheadingCell("Hourly Rate"));

        // Add both the main header and subheading table to the "Finance" cell
        Cell financeCell = new Cell(1, 4)
                .add(financeMainHeader)
                .add(subheadingTable)
                .setBackgroundColor(headerColor)
                .setFontColor(ColorConstants.RED)
                .setTextAlignment(TextAlignment.CENTER)
                .setVerticalAlignment(VerticalAlignment.MIDDLE);

        // Add the "Finance" cell with subheadings side by side to the table
        table.addCell(financeCell);

        table.addCell(createHeaderCell("Hours", headerColor));
        List<Map<String, String>> timesheets = (List<Map<String, String>>) sheetData.get("timesheets");

        // Iterate through each timesheet entry and populate the table
        for (Map<String, String> entry : timesheets) {
            // Add rows to the table with data from the JSON object
            table.addCell(createCell(entry.get("Work Package")));
            table.addCell(createCell(entry.get("Resource Name")));
            table.addCell(createCell(entry.get("Role")));
            table.addCell(createCell(entry.get("Location")));
            table.addCell(createCell(entry.get("Daily Rate")));
            table.addCell(createCell(entry.get("Hourly Rate")));
            table.addCell(createCell(entry.get("Hours")));
        }
        

        // Add the table to the document
        document.add(table);
        }


    // Method to add the other sheets with a uniform design
    private static void addSheetToPdf(Document document, Map<String, Object> sheetData) throws IOException {
    	 // Add project information from each sheet
        String projectName = (String) sheetData.get("projectName");
        String employeeName = (String) sheetData.get("EmployeeName");
        String poNumber = (String) sheetData.get("PO Number");

        // Load the image (logo) - replace with the correct path
        String logoPath = "C:\\Users\\hp\\Downloads\\Intelizign Logo\\image 3.png";
        ImageData imageData = ImageDataFactory.create(logoPath);
        Image logo = new Image(imageData);
        logo.setWidth(100); // Set appropriate size for the logo
        logo.setHeight(20);

        // Create header table with two columns
        Table headerTable = new Table(UnitValue.createPercentArray(new float[]{70, 30}));
        headerTable.setWidth(UnitValue.createPercentValue(100));
        headerTable.setFixedLayout(); // Set layout for the header table
        headerTable.setBorder(new SolidBorder(1f));
        headerTable.setHeight(UnitValue.createPercentValue(100));

        // Add the left-aligned header text cell
        Cell leftCell = new Cell()
                .add(new Paragraph("EmployeeName : "+employeeName)
                        .setBold()
                        .setFontSize(12)
                        .setTextAlignment(TextAlignment.LEFT))
                .add(new Paragraph("Po Number : "+poNumber) // Second paragraph
                        .setFontSize(10) // Adjust font size as needed
                        .setTextAlignment(TextAlignment.LEFT))
                .setBorder(null)
                .setPaddingRight(6)
                .setVerticalAlignment(VerticalAlignment.MIDDLE);
        headerTable.addCell(leftCell);

        // Add the right-aligned logo cell
        Cell rightCell = new Cell()
                .add(logo.setHorizontalAlignment(HorizontalAlignment.RIGHT))
                .setBorder(null)
                .setPaddingLeft(8)
                .setVerticalAlignment(VerticalAlignment.MIDDLE);
        headerTable.addCell(rightCell);

        // Add the header table to the document
        document.add(headerTable);
        Paragraph spacer = new Paragraph("").setMarginBottom(10); // Adjust the margin as needed
        document.add(spacer);
        
        
        // Create a table for timesheet data
        float[] columnWidths = {1, 2, 1, 2, 4};
        Table table = new Table(columnWidths);
        table.setWidth(UnitValue.createPercentValue(100));
        table.setHorizontalAlignment(HorizontalAlignment.CENTER);
        table.setBorder(new SolidBorder(1f));
        // Add table headers
        String[] headers = {"Slno", "Date", "Hours", "Work Package Name", "Activities"};
        for (String header : headers) {
            table.addHeaderCell(new Cell().add(new Paragraph(header).setFontSize(8).setBold().setTextAlignment(TextAlignment.CENTER)));
        }

        // Initialize total hours variable
        double totalHours = 0;

        // Populate table rows from sheet timesheet data
        List<Map<String, String>> timesheets = (List<Map<String, String>>) sheetData.get("timesheets");
        for (Map<String, String> entry : timesheets) {
            String bgColorHex = entry.get("backgroundColor");
            DeviceRgb bgColor = null;

            if (bgColorHex != null && !bgColorHex.isEmpty()) {
                int red = Integer.parseInt(bgColorHex.substring(1, 3), 16);
                int green = Integer.parseInt(bgColorHex.substring(3, 5), 16);
                int blue = Integer.parseInt(bgColorHex.substring(5, 7), 16);
                bgColor = new DeviceRgb(red, green, blue);
            }

            String hoursStr = entry.get("Hours");
            double hours = 0;
            try {
                hours = Double.parseDouble(hoursStr);
                totalHours += hours;
            } catch (NumberFormatException e) {
                System.out.println("Invalid hours format for entry: " + entry);
            }

            // Add each field in the entry to the table with optional background color
            addCellWithBackgroundColor(table, entry.get("Slno"));
            addCellWithBackgroundColor(table, entry.get("Date"));
            addCellWithBackgroundColor(table, entry.get("Hours"));
            addCellWithBackgroundColor(table, entry.get("work package Name"));
            addCellWithBackgroundColor(table, entry.get("Activities"));
        }

        // Add a row for the total hours
        table.addCell(new Cell(1, 2).add(new Paragraph("Total Hours").setFontSize(6)).setBold());
        table.addCell(new Cell(1, 3).add(new Paragraph(String.format("%.2f", totalHours)).setFontSize(6)).setBold());

        // Add the table to the document
        document.add(table);

    }
    // Helper method to add a cell with an optional background color
 // Helper method to add a cell without background color
    private static void addCellWithBackgroundColor(Table table, String content) {
        // Create a new cell and add the content
        Cell cell = new Cell().add(new Paragraph(content).setFontSize(6))
                .setTextAlignment(TextAlignment.CENTER)
                .setVerticalAlignment(VerticalAlignment.MIDDLE);

        // Allow text to wrap within the cell
        cell.setKeepTogether(false); // Allows content to flow to the next line if it's too long
        cell.setWordSpacing(0.5f); // Set word spacing to prevent cramped wrapping
        cell.setMinHeight(10); // Minimum height for the cell

        // Set padding and height of the cell to prevent cutting off data
        cell.setPadding(2);
        // cell.setBorder(new SolidBorder(1f));
        // cell.setHeightAuto(); // Ensure height adjusts automatically based on content

        // Add the cell to the table
        table.addCell(cell);
    }


 // Method to create a header cell with adjusted font size and cell padding
    private static Cell createHeaderCell(String text, Color backgroundColor) {
        return new Cell()
                .add(new Paragraph(text)) // Wrap the text in a Paragraph
                .setBackgroundColor(backgroundColor)
                .setFontColor(ColorConstants.RED)
                .setTextAlignment(TextAlignment.CENTER)
                .setVerticalAlignment(VerticalAlignment.MIDDLE)
                .setFontSize(9) // Adjusted font size
                .setPadding(5); // Adjust cell padding
    }


    // Method to create a data cell with adjusted font size and cell padding
    private static Cell createCell(String text) {
        return new Cell()
                .add(new Paragraph(text))
                .setFontSize(7)
                .setTextAlignment(TextAlignment.CENTER)
                .setVerticalAlignment(VerticalAlignment.MIDDLE)
                .setPadding(2);
    }
    
    private static Cell createSubheadingCell(String text) {
        return new Cell()
                .add(new Paragraph(text))
                .setFontColor(ColorConstants.RED)
                .setTextAlignment(TextAlignment.CENTER)
                .setVerticalAlignment(VerticalAlignment.MIDDLE)
                .setFontSize(7); // Adjust font size for subheadings
    }
}
