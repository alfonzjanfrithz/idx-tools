package com.example.idxdownloader;

import lombok.AllArgsConstructor;
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
import java.util.Map;

@Service
@AllArgsConstructor
public class ExcelDataWriterService {
    private static final int COL_KODE_EMITEN = 0;
    private static final int COL_PRICE = 1;
    private static final int COL_OUTSTANDING_SHARES = 2;
    private static final int COL_VOLUME = 3;
    private static final int COL_LIQUIDITY = 4;
    private static final int COL_TOTAL_LIABILITIES = 5;
    private static final int COL_TOTAL_EQUITIES = 6;
    private static final int COL_NET_PROFIT = 7;
    private static final int COL_NET_PROFIT_LAST_YEAR = 8;
    private static final int COL_MARKET_CAP = 9;
    private static final int COL_DER = 10;
    private static final int COL_PBV = 11;
    private static final int COL_PER = 12;
    private static final int COL_ROE = 13;
    private static final int COL_EPS = 14;
    private static final int COL_EPS_LAST_YEAR = 15;
    private static final int COL_LABA_NAIK_TURUN = 16;
    private static final int COL_ROUGH_EXPECTED_PRICE = 17;
    private static final int COL_MOS = 18;
    private static final int COL_TURNAROUND = 19;
    private static final int COL_DATE_ADDED = 20;


    public void updateOrCreateExcel(String year, String period, String kodeEmiten, FinancialData financialData, Map<String,TradingSummary> tradingSummary) throws IOException {
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
            sheet = workbook.createSheet(String.format("%s-%s", year, period));
            lastRowNum = 0;

            // Create header
            createHeader(workbook, sheet);
            lastRowNum++;

            // Copy "Data" sheet from shares_detail_20230929.xlsx to this workbook
            copyDataSheetToWorkbook(workbook);
        }

        // Add data
        addDataRow(sheet, lastRowNum, kodeEmiten, financialData, workbook, period, tradingSummary.get(kodeEmiten));

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
        headerStyle.setBorderTop(BorderStyle.MEDIUM);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);

        String[] headers = {
                "Kode Emiten",
                "Price (Rp)",
                "Shares",
                "Volume",
                "Liquidity",
                "Liabilities",
                "Equities",
                "Net Profit",
                "Net Profit L/Y",
                "Market Cap",
                "DER (x)",
                "PBV (x)",
                "PER (x)",
                "ROE (%)",
                "EPS (Rp)",
                "EPS L/Y (Rp)",
                "Laba (%)",
                "Rough Exp Price",
                "MoS (%)",
                "Turnaround",
                "Date Added"};

        for (int i = 0; i < headers.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(headers[i]);
            headerCell.setCellStyle(headerStyle);
        }
    }

    private void addDataRow(XSSFSheet sheet, int rowNum, String kodeEmiten, FinancialData financialData, XSSFWorkbook workbook, String period, TradingSummary tradingSummary) {
        double multiplier = getMultiplierForPeriod(period);
        Row row = sheet.createRow(rowNum);
        int currentRow = rowNum + 1;


        setCellValue(row, COL_KODE_EMITEN, kodeEmiten, getPlainStyle(workbook));
        setCellFormula(row, COL_PRICE, createFormula("VLOOKUP(A%d,'Data'!$B$2:$Y$2741,4,FALSE)", currentRow), getCurrencyStyle(workbook));
        setCellFormula(row, COL_OUTSTANDING_SHARES, createFormula("VLOOKUP(A%d,'Data'!$B$2:$Y$2741,3,FALSE)/1000000000", currentRow),  getDecimalStyle(workbook));
        if (tradingSummary != null) {
            setCellValue(row, COL_VOLUME, tradingSummary.getVolume(), getDecimalStyle(workbook));
        }
        setCellFormula(row, COL_LIQUIDITY, createFormula("%s%d*%s%d/1000000000", colIndexToLetter(COL_VOLUME), currentRow, colIndexToLetter(COL_PRICE), currentRow),  getDecimalStyle(workbook));
        setCellValue(row, COL_TOTAL_LIABILITIES, financialData.getTotalLiabilities(), getCurrencyStyle(workbook));
        setCellValue(row, COL_TOTAL_EQUITIES, financialData.getTotalEquities(), getCurrencyStyle(workbook));
        setCellValue(row, COL_NET_PROFIT, financialData.getNetProfit(), getCurrencyStyle(workbook));
        setCellValue(row, COL_NET_PROFIT_LAST_YEAR, financialData.getNetProfitLastYear(), getCurrencyStyle(workbook));
        setCellFormula(row, COL_MARKET_CAP, createFormula("%s%d*%s%d", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow), getCurrencyStyle(workbook));
        setCellFormula(row, COL_DER, createFormula("%s%d/%s%d", colIndexToLetter(COL_TOTAL_LIABILITIES), currentRow, colIndexToLetter(COL_TOTAL_EQUITIES), currentRow),  getDecimalStyle(workbook));
        setCellFormula(row, COL_PBV, createFormula("(%s%d*%s%d)/%s%d", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow, colIndexToLetter(COL_TOTAL_EQUITIES), currentRow),  getDecimalStyle(workbook));
        setCellFormula(row, COL_PER, createFormula("%s%d/%s%d", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_EPS), currentRow),  getDecimalStyle(workbook));
        setCellFormula(row, COL_ROE, createFormula("(%s%d/%s%d)*%f", colIndexToLetter(COL_NET_PROFIT), currentRow, colIndexToLetter(COL_TOTAL_EQUITIES), currentRow, multiplier), getPercentageStyle(workbook));
        setCellFormula(row, COL_EPS, createFormula("(%s%d/%s%d)*%f", colIndexToLetter(COL_NET_PROFIT), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow, multiplier),  getDecimalStyle(workbook));
        setCellFormula(row, COL_EPS_LAST_YEAR, createFormula("(%s%d/%s%d)*%f", colIndexToLetter(COL_NET_PROFIT_LAST_YEAR), currentRow, colIndexToLetter(COL_OUTSTANDING_SHARES), currentRow, multiplier),  getDecimalStyle(workbook));
        setCellFormula(row, COL_LABA_NAIK_TURUN, createFormula("((%s%d/%s%d)-1)", colIndexToLetter(COL_NET_PROFIT), currentRow, colIndexToLetter(COL_NET_PROFIT_LAST_YEAR), currentRow), getPercentageStyle(workbook));
        setCellFormula(row, COL_ROUGH_EXPECTED_PRICE, createFormula("((%s%d/%s%d)*(%s%d*10))", colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_PBV), currentRow, colIndexToLetter(COL_ROE), currentRow), getCurrencyStyle(workbook));

        if (financialData.getNetProfit() >= 0) {
            setCellFormula(row, COL_MOS, createFormula("(%s%d-%s%d)/%s%d", colIndexToLetter(COL_ROUGH_EXPECTED_PRICE), currentRow, colIndexToLetter(COL_PRICE), currentRow, colIndexToLetter(COL_ROUGH_EXPECTED_PRICE), currentRow), getPercentageStyle(workbook));
        } else {
            setCellValue(row, COL_MOS, "P. Loss", getPlainStyle(workbook));
        }

        if (financialData.getNetProfit() >= 0 && financialData.getNetProfitLastYear() <=0) {
            setCellValue(row, COL_TURNAROUND, "TA+", getPlainStyle(workbook));
        } else if (financialData.getNetProfit() <= 0 && financialData.getNetProfitLastYear() >=0) {
            setCellValue(row, COL_TURNAROUND, "TA-", getPlainStyle(workbook));
        } else {
            setCellValue(row, COL_TURNAROUND, "N/A", getPlainStyle(workbook));
        }

        setCellValue(row, COL_DATE_ADDED, LocalDate.now(), getDateStyle(workbook));

        applyNegativeValueRedFormatting(sheet, COL_PER);
        applyNegativeValueRedFormatting(sheet, COL_PBV);
        applyNegativeValueRedFormatting(sheet, COL_ROE);
        applyNegativeValueRedFormatting(sheet, COL_TOTAL_EQUITIES);
        applyNegativeValueRedFormatting(sheet, COL_NET_PROFIT);
        applyNegativeValueRedFormatting(sheet, COL_NET_PROFIT_LAST_YEAR);
        applyNegativeValueRedFormatting(sheet, COL_EPS);
        applyNegativeValueRedFormatting(sheet, COL_EPS_LAST_YEAR);
        applyNegativeValueRedFormatting(sheet, COL_MOS);

//        for (int i = 0; i <= COL_DATE_ADDED; i++) {
//            applyThickBorderToColumn(sheet, i, workbook);
//        }

        autoSizeColumn(sheet);
    }

    private void applyThickBorderToColumn(XSSFSheet sheet, int colIndex, XSSFWorkbook workbook) {

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                CellStyle currentStyle = cell.getCellStyle();
                CellStyle newStyle = workbook.createCellStyle();
                newStyle.cloneStyleFrom(currentStyle);

                // Add thick borders
                newStyle.setBorderLeft(BorderStyle.MEDIUM);
                newStyle.setBorderRight(BorderStyle.MEDIUM);

                cell.setCellStyle(newStyle);
            }
        }
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
        Cell cell = row.createCell(colIndex);
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
        }else if (value instanceof Long) {
            cell.setCellValue((Long) value);
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

    private CellStyle getDecimalStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "0.00");
    }

    private CellStyle getPercentageStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "0.00%");
    }

    private CellStyle getDateStyle(XSSFWorkbook workbook) {
        return findOrCreateCellStyleByDataFormat(workbook, "dd-MM-yyyy");
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
