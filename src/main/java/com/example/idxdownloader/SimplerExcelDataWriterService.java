package com.example.idxdownloader;

import com.google.common.collect.ImmutableList;
import lombok.AllArgsConstructor;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFontFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

@Service
@AllArgsConstructor
public class SimplerExcelDataWriterService {
    private static final int COL_NO = 0;
    private static final int COL_KODE_EMITEN = 1;
    private static final int COL_PRICE = 2;
    private static final int COL_PER = 3;
    private static final int COL_PBV = 4;
    private static final int COL_DER = 5;
    private static final int COL_ROE = 6;
    private static final int COL_OUTSTANDING_SHARES = 7;
    private static final int COL_TOTAL_LIABILITIES = 8;
    private static final int COL_TOTAL_EQUITIES = 9;
    private static final int COL_NET_PROFIT = 10;
    private static final int COL_NET_PROFIT_LAST_YEAR = 11;
    private static final int COL_EPS = 12;
    private static final int COL_MARKET_CAP = 13;
    private static final int COL_LABA_NAIK_TURUN = 14;
    private static final int COL_OTHERS = 15;


    public void updateOrCreateExcel(String year, String period, String kodeEmiten, FinancialData financialData, Map<String,TradingSummary> tradingSummary) throws IOException {
        String targetFilePath = System.getProperty("user.home") + "\\.idx-tmp\\kalkulator-saham-" + year + "-" + period + "-AutoGen.xlsx";
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
            sheet = workbook.createSheet(String.format("Kalkulator Value %s-%s", year, period));
            lastRowNum = 0;

            // Create header
            createHeader(workbook, sheet, year, period);
            lastRowNum++;
        }

        // Add data
        addDataRow(sheet, lastRowNum, kodeEmiten, financialData, workbook, period, tradingSummary.get(kodeEmiten));
        sheet.createFreezePane(0, 1);

        POIXMLProperties properties = workbook.getProperties();
        POIXMLProperties.CoreProperties coreProperties = properties.getCoreProperties();
        coreProperties.setCreator("Teguh Hidayat");

        try (FileOutputStream os = new FileOutputStream(targetFile)) {
            workbook.write(os);
        }
        workbook.close();
    }

    private void addOrUpdateRow(XSSFSheet sheet, XSSFWorkbook workbook, int rowIndex, String value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        setCellValue(row, COL_OTHERS, value, getPlainStyle(workbook));
    }

    public static String mapQuarter(String input) throws IllegalArgumentException {
        switch (input) {
            case "I":
                return "Q1";
            case "II":
                return "Q2";
            case "III":
                return "Q3";
            case "Tahunan":
                return "Q4";
            default:
                throw new IllegalArgumentException("Invalid input: " + input);
        }
    }

    private void createHeader(XSSFWorkbook workbook, XSSFSheet sheet, String year, String period) {
        String quarter = mapQuarter(period);
        List<String> headerInput = ImmutableList.of(
                "Stocks",
                "Price (Rp)",
                "Shares",
                "Liabilities",
                "Equity",
                String.format("Net Profit %s %s ",quarter, year),
                String.format("Net Profit %s %d", quarter, Integer.parseInt(year)-1));

        Row headerRow = sheet.createRow(0);
        Font blackFont = workbook.createFont();
        blackFont.setBold(true);

        Font blueFont = workbook.createFont();
        blueFont.setColor(IndexedColors.LIGHT_BLUE.getIndex());
        blueFont.setBold(true);

        CellStyle blackHeader = workbook.createCellStyle();
        blackHeader.setFont(blackFont);
        blackHeader.setBorderTop(BorderStyle.MEDIUM);
        blackHeader.setBorderBottom(BorderStyle.MEDIUM);
        blackHeader.setBorderRight(BorderStyle.MEDIUM);
        blackHeader.setBorderLeft(BorderStyle.MEDIUM);

        CellStyle blueHeader = workbook.createCellStyle();
        blueHeader.setFont(blueFont);
        blueHeader.setBorderTop(BorderStyle.MEDIUM);
        blueHeader.setBorderBottom(BorderStyle.MEDIUM);
        blueHeader.setBorderRight(BorderStyle.MEDIUM);
        blueHeader.setBorderLeft(BorderStyle.MEDIUM);

        String[] headers = {
                "No.",
                "Stocks",
                "Price (Rp)",
                "PER (x)",
                "PBV (x)",
                "DER (x)",
                "ROE (%)",
                "Shares",
                "Liabilities",
                "Equity",
                String.format("Net Profit %s %s ",quarter, year),
                String.format("Net Profit %s %d", quarter, Integer.parseInt(year)-1),
                "EPS (Rp)",
                "Market Cap",
                "Laba naik/turun?"};

        for (int i = 0; i < headers.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);

            if (headerInput.contains(headers[i])) {
                headerCell.setCellStyle(blueHeader);
            } else {
                headerCell.setCellStyle(blackHeader);
            }

        }
    }

    private void addDataRow(XSSFSheet sheet, int rowNum, String kodeEmiten, FinancialData financialData, XSSFWorkbook workbook, String period, TradingSummary tradingSummary) {
        double multiplier = getMultiplierForPeriod(period);
        Row row = sheet.getRow(rowNum) != null ? sheet.getRow(rowNum) : sheet.createRow(rowNum);
        int currentRow = rowNum + 1;
        Long price = 0L;
        Double listedShares = 0.0;

        if (tradingSummary != null) {
            price = tradingSummary.getClose();
            listedShares = Double.valueOf(tradingSummary.getListedShares()) /1000000000;
        } else {
            throw new RuntimeException(String.format("Cannot find tradingSummary for %s", kodeEmiten));
        }

        double pbv = price * listedShares / financialData.getTotalEquities();

        Long divider = 1L;
        if (pbv < 0.0001) {
            divider = 1_000_000L;
        }

        setCellValue(row, COL_NO, rowNum, getDecimalStyleNoComma(workbook));
        setCellValue(row, COL_KODE_EMITEN, kodeEmiten, getPlainStyle(workbook));
        setCellValue(row, COL_PRICE, price, getCurrencyStyle(workbook));
        setCellFormula(row, COL_PER, createFormula("%s%d/%s%d/%f", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_EPS), currentRow, multiplier),  getDecimalStyleBold(workbook));
        setCellFormula(row, COL_PBV, createFormula("(%s%d*%s%d)/%s%d", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow, colIndexToLetter(COL_TOTAL_EQUITIES), currentRow),  getDecimalStyleBold(workbook));
        setCellFormula(row, COL_DER, createFormula("%s%d/%s%d", colIndexToLetter(COL_TOTAL_LIABILITIES), currentRow, colIndexToLetter(COL_TOTAL_EQUITIES), currentRow),  getDecimalStyle(workbook));
        setCellFormula(row, COL_ROE, createFormula("(%s%d/%s%d)*%f", colIndexToLetter(COL_NET_PROFIT), currentRow, colIndexToLetter(COL_TOTAL_EQUITIES), currentRow, multiplier), getPercentageStyle(workbook));
        setCellValue(row, COL_OUTSTANDING_SHARES, listedShares, getDecimalStyle(workbook));
        setCellValue(row, COL_TOTAL_LIABILITIES, financialData.getTotalLiabilities()/divider, getCurrencyStyle(workbook));
        setCellValue(row, COL_TOTAL_EQUITIES, financialData.getTotalEquities()/divider, getCurrencyStyle(workbook));
        setCellValue(row, COL_NET_PROFIT, financialData.getNetProfit()/divider, getCurrencyStyle(workbook));
        setCellValue(row, COL_NET_PROFIT_LAST_YEAR, financialData.getNetProfitLastYear()/divider, getCurrencyStyle(workbook));
        setCellFormula(row, COL_EPS, createFormula("(%s%d/%s%d)", colIndexToLetter(COL_NET_PROFIT), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow),  getDecimalStyle(workbook));
        setCellFormula(row, COL_MARKET_CAP, createFormula("%s%d*%s%d", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow), getCurrencyStyle(workbook));
        setCellFormula(row, COL_LABA_NAIK_TURUN, createFormula("((%s%d/%s%d)-1)", colIndexToLetter(COL_NET_PROFIT), currentRow, colIndexToLetter(COL_NET_PROFIT_LAST_YEAR), currentRow), getPercentageStyle(workbook));

        applyNegativeValueRedFormatting(sheet, COL_PER);
        applyNegativeValueRedFormatting(sheet, COL_PBV);
        applyNegativeValueRedFormatting(sheet, COL_ROE);
        applyNegativeValueRedFormatting(sheet, COL_TOTAL_EQUITIES);
        applyNegativeValueRedFormatting(sheet, COL_NET_PROFIT);
        applyNegativeValueRedFormatting(sheet, COL_NET_PROFIT_LAST_YEAR);
        applyNegativeValueRedFormatting(sheet, COL_EPS);
        applyNegativeValueRedFormatting(sheet, COL_LABA_NAIK_TURUN);


        autoSizeColumn(sheet);
    }

    private void applyNegativeValueRedFormatting(XSSFSheet sheet, int colIndex) {
        XSSFWorkbook workbook = sheet.getWorkbook();
        XSSFSheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();

        XSSFFont font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());

        XSSFConditionalFormattingRule rule = conditionalFormatting.createConditionalFormattingRule(ComparisonOperator.LT, "0");
        XSSFFontFormatting fontFmt = rule.createFontFormatting();
        fontFmt.setFontColorIndex(IndexedColors.RED.index);

        CellRangeAddress[] regions = {
                new CellRangeAddress(1, sheet.getLastRowNum(), colIndex, colIndex)
        };

        conditionalFormatting.addConditionalFormatting(regions, rule);
    }


    private static void autoSizeColumn(XSSFSheet sheet) {
        int numberOfColumns = sheet.getRow(0).getLastCellNum();
        for (int i = 0; i < numberOfColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void setCellValue(Row row, int colIndex, Object value, CellStyle style) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
        }else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        }else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private void setCellFormula(Row row, int colIndex, String formula, CellStyle style) {
        Cell cell = row.createCell(colIndex);
        cell.setCellFormula(formula);
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private String createFormula(String pattern, Object... args) {
        return String.format(pattern, args);
    }

    private double getMultiplierForPeriod(String period) {
        switch (period) {
            case "I":
                return 4.0;
            case "II":
                return 2.0;
            case "III":
                return 4.0 / 3.0;
            case "Tahunan": // Assuming you meant four 'I's for the fourth quarter
                return 1.0;
            default:
                throw new IllegalArgumentException("Invalid period: " + period);
        }
    }


    private CellStyle findOrCreateCellStyleByDataFormat(XSSFWorkbook workbook, String dataFormat) {
        int numStyles = workbook.getNumCellStyles();

        for (int i = 0; i < numStyles; i++) {
            CellStyle existingStyle = workbook.getCellStyleAt(i);

            if (existingStyle.getDataFormatString().equalsIgnoreCase(dataFormat)) {
                return existingStyle;
            }
        }

        // If no matching style is found, create a new one and return it
        CellStyle style = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat(dataFormat));
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        return style;
    }

    private CellStyle getCurrencyStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "#,##0");
    }


    private CellStyle getDecimalStyleBold(XSSFWorkbook workbook) {
        CellStyle decimalStyle = findOrCreateCellStyleByDataFormat(workbook, "0.0");
        Font blackFont = workbook.createFont();
        blackFont.setBold(true);
        decimalStyle.setFont(blackFont);
        return decimalStyle;
    }

    private CellStyle getDecimalStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "0.00");
    }

    private CellStyle getDecimalStyleNoComma(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "0");
    }

    private CellStyle getPercentageStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "0.0%");
    }

    private CellStyle getPlainStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "General");
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
