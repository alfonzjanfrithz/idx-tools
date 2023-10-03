package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
@AllArgsConstructor
public class ExcelDataWriterService {
    public void updateOrCreateExcel(String year, String period, String kodeEmiten, FinancialData financialData) throws IOException, FileNotFoundException {
        String targetFilePath = System.getProperty("user.home") + "\\.idx-tmp\\Summary-" + year + "-" + period + ".xlsx";
        File targetFile = new File(targetFilePath);

        XSSFWorkbook workbook;
        XSSFSheet sheet;
        int lastRowNum;

        // Check if the file exists
        if (targetFile.exists()) {
            workbook = new XSSFWorkbook(new FileInputStream(targetFile));
            sheet = workbook.getSheetAt(0);
            lastRowNum = sheet.getLastRowNum() + 1;
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Sheet 1");
            lastRowNum = 0;

            // Create header
            createHeader(workbook, sheet);
            lastRowNum++;
        }

        // Add data
        addDataRow(sheet, lastRowNum, kodeEmiten, financialData, workbook, period);

        try (FileOutputStream os = new FileOutputStream(targetFile)) {
            workbook.write(os);
        }
        workbook.close();
    }

    private void createHeader(XSSFWorkbook workbook, XSSFSheet sheet) {
        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        String[] headers = {"Kode Emiten", "Total Liabilities", "Total Equities", "Net Profit", "Net Profit Last Year", "DER", "ROE"};
        for (int i = 0; i < headers.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);
            headerCell.setCellStyle(headerStyle);
        }
    }

    private void addDataRow(XSSFSheet sheet, int rowNum, String kodeEmiten, FinancialData financialData, XSSFWorkbook workbook, String period) {
        // Create styles
        CellStyle currencyStyle = getCurrencyStyle(workbook);
        CellStyle decimalStyle = getDecimalStyle(workbook);
        CellStyle percentageStyle = getPercentageStyle(workbook);

        // Add data
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(kodeEmiten);

        Cell totalLiabilitiesCell = row.createCell(1);
        double totalLiabilities = financialData.getTotalLiabilities();
        totalLiabilitiesCell.setCellValue(totalLiabilities);
        totalLiabilitiesCell.setCellStyle(currencyStyle);

        Cell totalEquitiesCell = row.createCell(2);
        double totalEquities  = financialData.getTotalEquities();
        totalEquitiesCell.setCellValue(totalEquities);
        totalEquitiesCell.setCellStyle(currencyStyle);

        Cell netProfitCell = row.createCell(3);
        double netProfit = financialData.getNetProfit();
        netProfitCell.setCellValue(netProfit);
        netProfitCell.setCellStyle(currencyStyle);

        Cell netProfitLastYearCell = row.createCell(4);
        double netProfitLastYear = financialData.getNetProfitLastYear();
        netProfitLastYearCell.setCellValue(netProfitLastYear);
        netProfitLastYearCell.setCellStyle(currencyStyle);

        Cell derCell = row.createCell(5);
        derCell.setCellValue(totalLiabilities / totalEquities);
        derCell.setCellStyle(decimalStyle);


        // Adjust ROE based on the period
        double adjustedNetProfit = netProfit;
        switch (period) {
            case "I":
                adjustedNetProfit = netProfit * 4;
                break;
            case "II":
                adjustedNetProfit = netProfit * 2;
                break;
            case "III":
                adjustedNetProfit = netProfit * (4.0 / 3.0);
                break;
            case "IIII": // Assuming you meant four 'I's for the fourth quarter
                // No adjustment needed
                break;
            default:
                throw new IllegalArgumentException("Invalid period: " + period);
        }

        Cell roeCell = row.createCell(6);
        roeCell.setCellValue(adjustedNetProfit / totalEquities);
        roeCell.setCellStyle(percentageStyle);

        int numberOfColumns = sheet.getRow(0).getLastCellNum();

        // Auto-size columns to fit content
        for (int i = 0; i < numberOfColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private CellStyle getCurrencyStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("#,##0"));
        return style;
    }

    private CellStyle getDecimalStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("0.00"));
        return style;
    }

    private CellStyle getPercentageStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("0.00%"));
        return style;
    }
}
