package com.example.idxdownloader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelDataReaderService {
    public static final int SH_GENERAL_INFO_IDX = 2;
    public static final int SH_BALANCE_SHEET = 3;
    public static final int SH_PROFIT_LOSS = 4;

    public static final String KEY_MATA_UANG_PELAPORAN = "Mata uang pelaporan";
    public static final String KEY_PEMBULATAN = "Pembulatan yang digunakan dalam penyajian jumlah dalam laporan keuangan";
    public static final String CURR_LOV_RUPIAH_IDR = "Rupiah / IDR";
    public static final String KEY_JUMLAH_LIABILITAS = "Jumlah liabilitas";
    public static final String KEY_EKUITAS_ENTITAS_INDUK = "Jumlah ekuitas yang diatribusikan kepada pemilik entitas induk";
    public static final String KEY_LABA_RUGI_ENTITAS_INDUK = "Laba (rugi) yang dapat diatribusikan ke entitas induk";
    public static final String KEY_PENJUALAN_DAN_PENDAPATAN_USAHA = "Penjualan dan pendapatan usaha";
    public static final String KEY_BEBAN_POKOK = "Beban pokok penjualan dan pendapatan";

    public FinancialData readFinancialData(String filePath) throws IOException, InvalidFormatException {
        FinancialData financialData = new FinancialData();

        final double USD_CONVERSION_RATE = 15000.0; // Conversion rate for USD to IDR
        final double BILLION = 1_000_000_000.0; // 1 billion

        ZipSecureFile.setMinInflateRatio(0.002);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new File(filePath))) {
            XSSFSheet info = workbook.getSheetAt(SH_GENERAL_INFO_IDX);
            boolean isIDRCurrency = isInIDR(info);
            long multiplier = multiplierValue(info);

            XSSFSheet balanceSheet = workbook.getSheetAt(SH_BALANCE_SHEET);
            double totalLiabilities = getTotalLiabilities(balanceSheet);
            double totalEquities = getTotalEquities(balanceSheet);

            XSSFSheet profitLoss = workbook.getSheetAt(SH_PROFIT_LOSS);
            double netProfit = getNetProfit(profitLoss);
            double netProfitLastYear = getNetProfitLastYear(profitLoss);
            double revenue = getRevenue(profitLoss);

            // Convert values if they are in USD
            if (!isIDRCurrency) {
                totalLiabilities *= USD_CONVERSION_RATE;
                totalEquities *= USD_CONVERSION_RATE;
                revenue *= USD_CONVERSION_RATE;
                netProfit *= USD_CONVERSION_RATE;
                netProfitLastYear *= USD_CONVERSION_RATE;
            }

            // Normalize values based on the multiplier and convert to billions
            totalLiabilities = (totalLiabilities * multiplier) / BILLION;
            totalEquities = (totalEquities * multiplier) / BILLION;
            revenue = (revenue * multiplier) / BILLION;
            netProfit = (netProfit * multiplier) / BILLION;
            netProfitLastYear = (netProfitLastYear * multiplier) / BILLION;

            // Set the normalized values to the financialData object
            financialData.setTotalLiabilities(totalLiabilities);
            financialData.setTotalEquities(totalEquities);
            financialData.setNetProfit(netProfit);
            financialData.setNetProfitLastYear(netProfitLastYear);
            financialData.setRevenue(revenue);
            financialData.setIDRCurrency(isIDRCurrency);
            financialData.setMultiplier(multiplier);
        }

        return financialData;
    }

    private static Double getTotalLiabilities(XSSFSheet sheet) throws IOException {
        String jumlahLiabilitias = findRowValue(KEY_JUMLAH_LIABILITAS, sheet);
        return Double.parseDouble(jumlahLiabilitias);
    }

    private static Double getTotalEquities(XSSFSheet sheet) throws IOException {
        String totalEquities = findRowValue(KEY_EKUITAS_ENTITAS_INDUK, sheet);
        return Double.parseDouble(totalEquities);
    }

    private static Double getNetProfit(XSSFSheet sheet) throws IOException {
        String netProfit = findRowValue(KEY_LABA_RUGI_ENTITAS_INDUK, sheet);
        return Double.parseDouble(netProfit);
    }

    private static Double getRevenue(XSSFSheet sheet) throws IOException {
        String netProfit = findRowValue(KEY_PENJUALAN_DAN_PENDAPATAN_USAHA, sheet);
        return Double.parseDouble(netProfit);
    }

    private static Double getNetProfitLastYear(XSSFSheet sheet) throws IOException {
        String netProfitLastYear = findRowValue(KEY_LABA_RUGI_ENTITAS_INDUK, sheet, 2);
        return Double.parseDouble(netProfitLastYear);
    }

    private static boolean isInIDR(XSSFSheet sheet) throws IOException {
        String currency = findRowValue(KEY_MATA_UANG_PELAPORAN, sheet);
        return CURR_LOV_RUPIAH_IDR.equalsIgnoreCase(currency);
    }

    private static long multiplierValue(XSSFSheet sheet) throws IOException {
        String multiplier = findRowValue(KEY_PEMBULATAN, sheet);

        return switch (multiplier) {
            case "Satuan Penuh / Full Amount" -> 1L;
            case "Ribuan / In Thousand" -> 1_000L;
            case "Jutaan / In Million" -> 1_000_000L;
            case "Miliaran / In Billion" -> 1_000_000_000L;
            default -> throw new IllegalArgumentException("Invalid multiplier: " + multiplier);
        };
    }

    private static String findRowValue(String fieldName, XSSFSheet sheetInput) throws IOException {
        return findRowValue(fieldName, sheetInput, 1);
    }

    private static String findRowValue(String fieldName, XSSFSheet sheetInput, int columnNo) throws IOException {
        List<List<String>> rows = getAllRows(sheetInput);

        return rows.stream()
                .filter(r -> fieldName.equalsIgnoreCase(r.getFirst()))
                .map(r -> r.get(columnNo)).findFirst().orElseThrow();
    }

    private static List<List<String>> getAllRows(XSSFSheet sheet) throws IOException {
        List<List<String>> rows = new ArrayList<>();

        for (Row row : sheet) {
            List<String> rowData = new ArrayList<>();
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case CellType.NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            rowData.add(cell.getDateCellValue().toString()); // if it's a date
                        } else {
                            double cellValue = cell.getNumericCellValue();
                            if (cellValue % 1 == 0) { // Check if it's an integer
                                rowData.add(String.valueOf((long) cellValue));
                            } else {
                                rowData.add(String.valueOf(cellValue));
                            }
                        }
                        break;
                    case CellType.STRING:
                        rowData.add(cell.getStringCellValue());
                        break;
                    case CellType.BOOLEAN:
                        rowData.add(String.valueOf(cell.getBooleanCellValue()));
                        break;
                    case CellType.FORMULA:
                        // Handle formula cells if needed
                        rowData.add(cell.getCellFormula());
                        break;
                    default:
                        rowData.add(""); // or handle other cell types if needed
                        break;
                }
            }
            rows.add(rowData);
        }
        return rows;
    }
}