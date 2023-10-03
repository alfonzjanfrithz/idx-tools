package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.stereotype.Service;

import java.io.IOException;

@Service
@AllArgsConstructor
public class ExcelReaderService {
    private FileDownloadService fileService;
    private final ExcelDataReaderService dataReaderService;
    private final ExcelDataWriterService dataWriterService;


    public void readExcel(String year, String period, String kodeEmiten) throws IOException, InvalidFormatException {
        String filePath = fileService.getFilePath(year, period, kodeEmiten);

        if (!fileService.fileExists(filePath)) {
            fileService.downloadFS(Integer.parseInt(year), period, kodeEmiten);
        }

        FinancialData financialData = dataReaderService.readFinancialData(filePath);
        dataWriterService.updateOrCreateExcel(year, period, kodeEmiten, financialData);
    }
}