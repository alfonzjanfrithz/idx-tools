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
    private static final int COL_KODE_EMITEN = 0;
    private static final int COL_PRICE = 1;
    private static final int COL_OUTSTANDING_SHARES = 2;
    private static final int COL_TOTAL_LIABILITIES = 3;
    private static final int COL_TOTAL_EQUITIES = 4;
    private static final int COL_NET_PROFIT = 5;
    private static final int COL_NET_PROFIT_LAST_YEAR = 6;
    private static final int COL_DER = 7;
    private static final int COL_ROE = 8;
    private static final int COL_EPS = 9;
    private static final int COL_EPS_LAST_YEAR = 10;
    private static final int COL_LABA_NAIK_TURUN = 11;

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

        String[] headers = {
                "Kode Emiten",
                "Price (Rp)",
                "Outstanding Shares",
                "Total Liabilities",
                "Total Equities",
                "Net Profit",
                "Net Profit Last Year",
                "DER",
                "ROE (%)",
                "EPS (Earning/Share)",
                "EPS Last Year",
                "Laba Naik/Turun"};

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
        row.createCell(COL_KODE_EMITEN).setCellValue(kodeEmiten);

        Cell priceCell = row.createCell(COL_PRICE);
        String priceFormula = "VLOOKUP(A" + (rowNum+1) + ",'Data'!$B$2:$Y$2741,4,FALSE)";
        priceCell.setCellFormula(priceFormula);
        priceCell.setCellStyle(currencyStyle);

        Cell outstandingSharesCell = row.createCell(COL_OUTSTANDING_SHARES);
        String outstandingShareFormula = "VLOOKUP(A" + (rowNum+1) + ",'Data'!$B$2:$Y$2741,3,FALSE)/1000000000";
        outstandingSharesCell.setCellFormula(outstandingShareFormula);
        outstandingSharesCell.setCellStyle(decimalStyle);

        Cell totalLiabilitiesCell = row.createCell(COL_TOTAL_LIABILITIES);
        double totalLiabilities = financialData.getTotalLiabilities();
        totalLiabilitiesCell.setCellValue(totalLiabilities);
        totalLiabilitiesCell.setCellStyle(currencyStyle);

        Cell totalEquitiesCell = row.createCell(COL_TOTAL_EQUITIES);
        double totalEquities  = financialData.getTotalEquities();
        totalEquitiesCell.setCellValue(totalEquities);
        totalEquitiesCell.setCellStyle(currencyStyle);

        Cell netProfitCell = row.createCell(COL_NET_PROFIT);
        double netProfit = financialData.getNetProfit();
        netProfitCell.setCellValue(netProfit);
        netProfitCell.setCellStyle(currencyStyle);

        Cell netProfitLastYearCell = row.createCell(COL_NET_PROFIT_LAST_YEAR);
        double netProfitLastYear = financialData.getNetProfitLastYear();
        netProfitLastYearCell.setCellValue(netProfitLastYear);
        netProfitLastYearCell.setCellStyle(currencyStyle);

        Cell derCell = row.createCell(COL_DER);
        String totalLiabilitiesCol = colIndexToLetter(COL_TOTAL_LIABILITIES);
        String totalEquitiesCol = colIndexToLetter(COL_TOTAL_EQUITIES);
        String derFormula = totalLiabilitiesCol + (rowNum+1) + "/" + totalEquitiesCol + (rowNum+1);

        derCell.setCellFormula(derFormula);
        derCell.setCellStyle(decimalStyle);

        // Adjust ROE based on the period
        double multiplier = getMultiplierForPeriod(period);
        String netProfitCol = colIndexToLetter(COL_NET_PROFIT);
        String roeFormulaBase = "(" + netProfitCol + (rowNum+1) + "/" + totalEquitiesCol + (rowNum+1) + ")*" + multiplier;
        Cell roeCell = row.createCell(COL_ROE);
        roeCell.setCellFormula(roeFormulaBase);
        roeCell.setCellStyle(percentageStyle);

        int numberOfColumns = sheet.getRow(0).getLastCellNum();

        // Auto-size columns to fit content
        for (int i = 0; i < numberOfColumns; i++) {
            sheet.autoSizeColumn(i);
        }

        Cell epsCell = row.createCell(COL_EPS);
        String outstandingSharesCol = colIndexToLetter(COL_OUTSTANDING_SHARES);
        String epsFormula =  "(" + netProfitCol + (rowNum+1) + "/" + outstandingSharesCol + (rowNum+1) + ")*" + multiplier;;

        epsCell.setCellFormula(epsFormula);
        epsCell.setCellStyle(decimalStyle);

        Cell epsLastyearCell = row.createCell(COL_EPS_LAST_YEAR);
        String epsLastYearFormula =  "(" + netProfitLastYear + (rowNum+1) + "/" + outstandingSharesCol + (rowNum+1) + ")*" + multiplier;;
        epsLastyearCell.setCellFormula(epsLastYearFormula);
        epsLastyearCell.setCellStyle(decimalStyle);

        Cell labaNaikCell = row.createCell(COL_LABA_NAIK_TURUN);
        String netProfitLastYearCol = colIndexToLetter(COL_NET_PROFIT_LAST_YEAR);
        String labaNaikTurunFormula = "(("+ netProfitCol + (rowNum+1) + "/" + netProfitLastYearCol + (rowNum+1) +")-1)";
        labaNaikCell.setCellFormula(labaNaikTurunFormula);
        labaNaikCell.setCellStyle(percentageStyle);
    }

    private double getMultiplierForPeriod(String period) {
        switch (period) {
            case "I":
                return 4.0;
            case "II":
                return 2.0;
            case "III":
                return 4.0 / 3.0;
            case "IIII": // Assuming you meant four 'I's for the fourth quarter
                return 1.0;
            default:
                throw new IllegalArgumentException("Invalid period: " + period);
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

    private static String colIndexToLetter(int colIndex) {
        StringBuilder columnName = new StringBuilder();
        while (colIndex >= 0) {
            int currentChar = colIndex % 26 + 'A';
            columnName.append((char) currentChar);
            colIndex = colIndex / 26 - 1;
        }
        return columnName.reverse().toString();
    }
}
