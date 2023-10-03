package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.shell.command.annotation.Command;

import java.io.IOException;


@Command
@AllArgsConstructor
public class FinancialDataCommands {
    private FileDownloadService fileDownloadService;
    private ExcelReaderService excelReaderService;

    // downloadFinancialStatement 2023 II ANJT
    @Command(command = "downloadFinancialStatement", description = "Download the financial statement for the given year, period, and kodeEmiten.")
    public void downloadFinancialStatement(int year, String period, String kodeEmiten) {
        fileDownloadService.downloadFS(year, period, kodeEmiten);
    }

    // readFinancialData 2023 II HRTA,ANJT,KEJU,SBMA,PURA,KLBF,BAYU,BSML,INCO,MITI,ADMF
    @Command(command = "readFinancialData", description = "Read the financial data for the given year, period, and multiple kodeEmiten values.")
    public void readFinancialData(String year, String period, String kodeEmitenList) {
        String[] kodeEmitens = kodeEmitenList.split(",");
        for (String kodeEmiten : kodeEmitens) {
            try {
                excelReaderService.readExcel(year, period, kodeEmiten.trim());
            } catch (IOException e) {
                e.printStackTrace();
            } catch (InvalidFormatException e) {
                throw new RuntimeException(e);
            }
        }
    }

    // TODO: Add PER, PBV, EPS, Market Cap, Laba Naik/Turun
    // TODO; Harga saham

}