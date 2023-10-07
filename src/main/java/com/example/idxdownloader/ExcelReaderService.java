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


    public void readExcel(String year, String period, String kodeEmiten, Map<String, TradingSummary> tradingSummary) throws IOException, InvalidFormatException {
        FinancialData financialData = getFinancialData(year, period, kodeEmiten);
        dataWriterService.updateOrCreateExcel(year, period, kodeEmiten, financialData, tradingSummary);
    }

    public FinancialData getFinancialData(String year, String period, String kodeEmiten) throws IOException, InvalidFormatException {
        String filePath = fileService.getFilePath(year, period, kodeEmiten);

        if (!fileService.fileExists(filePath)) {
            fileService.downloadFS(Integer.parseInt(year), period, kodeEmiten);
        }

        FinancialData financialData = dataReaderService.readFinancialData(filePath);
        return financialData;
    }
}