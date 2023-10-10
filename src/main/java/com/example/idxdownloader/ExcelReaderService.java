package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.Map;

@Service
@AllArgsConstructor
public class ExcelReaderService {
    private final FileDownloadService fileService;
    private final ExcelDataReaderService dataReaderService;
    private final ExcelDataWriterService dataWriterService;
    private final SimplerExcelDataWriterService simplerDataWriterService;


    public final Long DEFAULT_USD_CONVERSION_RATE = 15000L; // Conversion rate for USD to IDR


    public void readExcel(String year, String period, String kodeEmiten, Map<String, TradingSummary> tradingSummary) throws IOException, InvalidFormatException {
        FinancialData financialData = getFinancialData(year, period, kodeEmiten);
        dataWriterService.updateOrCreateExcel(year, period, kodeEmiten, financialData, tradingSummary);
    }

    public void simplerReadExcel(String year, String period, String kodeEmiten, Long usdIdrRate) throws IOException, InvalidFormatException {
        FinancialData financialData = getFinancialData(year, period, kodeEmiten, usdIdrRate);
        simplerDataWriterService.updateOrCreateExcel(year, period, kodeEmiten, financialData);
    }

    public FinancialData getFinancialData(String year, String period, String kodeEmiten) throws IOException, InvalidFormatException {
        return getFinancialData(year, period, kodeEmiten, DEFAULT_USD_CONVERSION_RATE);
    }

    public FinancialData getFinancialData(String year, String period, String kodeEmiten, Long usdIdrRate) throws IOException, InvalidFormatException {
        if (usdIdrRate == null) {
            usdIdrRate = DEFAULT_USD_CONVERSION_RATE;
        }
        String filePath = fileService.getFilePath(year, period, kodeEmiten);

        if (!fileService.fileExists(filePath)) {
            fileService.downloadFS(Integer.parseInt(year), period, kodeEmiten);
        }

        FinancialData financialData = dataReaderService.readFinancialData(filePath, usdIdrRate);
        return financialData;
    }
}