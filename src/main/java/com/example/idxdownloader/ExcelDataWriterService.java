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

            // Copy "Data" sheet from shares_detail_20230929.xlsx to this workbook
            copyDataSheetToWorkbook(workbook);
        }

        // Add data
        addDataRow(sheet, lastRowNum, kodeEmiten, financialData, workbook, period);

        try (FileOutputStream os = new FileOutputStream(targetFile)) {
            workbook.write(os);
        }
        workbook.close();
    }

    private void copyDataSheetToWorkbook(XSSFWorkbook targetWorkbook) throws IOException {
        String sourceFilePath = "shares_detail_20230929.xlsx"; // Assuming it's in the root folder
        FileInputStream fis = new FileInputStream(sourceFilePath);
        XSSFWorkbook sourceWorkbook = new XSSFWorkbook(fis);
        XSSFSheet sourceSheet = sourceWorkbook.getSheet("Data");

        XSSFSheet targetSheet = targetWorkbook.createSheet("Data");

        for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row targetRow = targetSheet.createRow(i);
            if (sourceRow != null) {
                for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceRow.getCell(j);
                    Cell targetCell = targetRow.createCell(j);
                    if (sourceCell != null) {
                        targetCell.setCellValue(sourceCell.toString());
                    }
                }
            }
        }

        sourceWorkbook.close();
        fis.close();
    }

    private void createHeader(XSSFWorkbook workbook, XSSFSheet sheet) {
        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);

        String[] headers = {"Kode Emiten", "Outstanding Shares", "Total Liabilities", "Total Equities", "Net Profit", "Net Profit Last Year", "DER", "ROE"};
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

        Cell outstandingSharesCell = row.createCell(1);
        String formula = "VLOOKUP(A" + (rowNum+1) + ",'Data'!$B$2:$Y$2741,3,FALSE)/1000000000";
        outstandingSharesCell.setCellFormula(formula);
        outstandingSharesCell.setCellStyle(decimalStyle);

        Cell totalLiabilitiesCell = row.createCell(2);
        double totalLiabilities = financialData.getTotalLiabilities();
        totalLiabilitiesCell.setCellValue(totalLiabilities);
        totalLiabilitiesCell.setCellStyle(currencyStyle);

        Cell totalEquitiesCell = row.createCell(3);
        double totalEquities  = financialData.getTotalEquities();
        totalEquitiesCell.setCellValue(totalEquities);
        totalEquitiesCell.setCellStyle(currencyStyle);

        Cell netProfitCell = row.createCell(4);
        double netProfit = financialData.getNetProfit();
        netProfitCell.setCellValue(netProfit);
        netProfitCell.setCellStyle(currencyStyle);

        Cell netProfitLastYearCell = row.createCell(5);
        double netProfitLastYear = financialData.getNetProfitLastYear();
        netProfitLastYearCell.setCellValue(netProfitLastYear);
        netProfitLastYearCell.setCellStyle(currencyStyle);

        Cell derCell = row.createCell(6);
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

        Cell roeCell = row.createCell(7);
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
