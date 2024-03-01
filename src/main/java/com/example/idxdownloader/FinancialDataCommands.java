package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.shell.command.annotation.Command;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;


@Command
@AllArgsConstructor
public class FinancialDataCommands {
    public static final String ANSI_RESET = "\u001B[0m";
    public static final String ANSI_RED = "\u001B[31m";
    public static final String ANSI_GREEN = "\u001B[32m";
    public static final String ANSI_CYAN = "\u001B[36m";
    public static final String ANSI_YELLOW = "\u001B[33m";
    public static final String ANSI_BLUE = "\u001B[34m";

    private ExcelReaderService excelReaderService;
    private TradingSummaryService tradingSummaryService;
    private ValueSheetService valueSheetService;

    // valueSheet 2023 II ANJT
    @Command(command = "valueSheet", description = "Populate value sheet")
    public void valueSheet(String kodeEmiten) throws IOException, InvalidFormatException {
        valueSheetService.populateTemplate(kodeEmiten);
    }

    // readFinancialData 2023 II HRTA,ANJT,GOTO,KEJU,SBMA,PURA,KLBF,BAYU,BSML,INCO,MITI,ADMF
    @Command(command = "readFinancialData", description = "Read the financial data for the given year, period, and multiple kodeEmiten values.")
    public void readFinancialData(String year, String period, String kodeEmitenList) {
        String[] kodeEmitens = kodeEmitenList.split(",");
        int currentCount = 0;
        int successfulCount = 0;
        int failedCount = 0;
        Map<String, String> failedDetails = new HashMap<>();
        long totalProcessingTime = 0;

        long startTimeOverall = System.currentTimeMillis(); // Start time for the entire process
        Map<String, TradingSummary> tradingSummary = tradingSummaryService.getTradingSummary();
        for (String kodeEmiten : kodeEmitens) {
            currentCount++;
            long startTime = System.currentTimeMillis(); // Start time for this kodeEmiten

            System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
            System.out.println(ANSI_CYAN + "🔄 Processing data for kodeEmiten: " + kodeEmiten.trim() +" ("+ currentCount + "/" + kodeEmitens.length + ") ..." + ANSI_RESET);
            try {
                excelReaderService.readExcel(year, period, kodeEmiten.trim(), tradingSummary);
                System.out.println(ANSI_GREEN + "✅ Successfully processed data for kodeEmiten: " + kodeEmiten.trim() + ANSI_RESET);
                successfulCount++;
            } catch (Exception e) {
                String errorMessage = e.getMessage();
                failedDetails.put(kodeEmiten.trim(), errorMessage);
                failedCount++;

                if (e instanceof IOException) {
                    System.out.println(ANSI_RED + "❌ Error processing data for kodeEmiten: " + kodeEmiten.trim() + ". Reason: " + errorMessage + ANSI_RESET);
                } else if (e instanceof InvalidFormatException) {
                    System.out.println(ANSI_RED + "⚠️ Invalid format encountered for kodeEmiten: " + kodeEmiten.trim() + ". Reason: " + errorMessage + ANSI_RESET);
                } else {
                    System.out.println(ANSI_RED + "❌ Unknown Error processing data for kodeEmiten: " + kodeEmiten.trim() + ". Reason: " + errorMessage + ANSI_RESET);
                }
            }
            long endTime = System.currentTimeMillis(); // End time for this kodeEmiten
            totalProcessingTime += (endTime - startTime);
            System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
        }

        // Summary
        System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
        System.out.println(ANSI_CYAN + "📊 SUMMARY 📊" + ANSI_RESET);
        System.out.println(ANSI_GREEN + "✅ Total Successful: " + successfulCount + ANSI_RESET);
        System.out.println(ANSI_RED + "❌ Total Failed: " + failedCount + ANSI_RESET);
        System.out.println(ANSI_BLUE + "📁 Total Files Processed: " + kodeEmitens.length + ANSI_RESET);

        if (failedCount > 0) {
            System.out.println(ANSI_RED + "\n❌ Failed Details:" + ANSI_RESET);
            for (Map.Entry<String, String> entry : failedDetails.entrySet()) {
                System.out.println(ANSI_RED + "❌ " + entry.getKey() + ": " + entry.getValue() + ANSI_RESET);
            }
        }

        long endTimeOverall = System.currentTimeMillis(); // End time for the entire process
        long totalTimeTaken = endTimeOverall - startTimeOverall;
        double averageTimePerKodeEmiten = (double) totalProcessingTime / kodeEmitens.length;


        System.out.println(ANSI_BLUE + "⏱️ Average Time Per KodeEmiten: " + formatTime((long) averageTimePerKodeEmiten) + ANSI_RESET);
        System.out.println(ANSI_BLUE + "⏱️ Total Time Taken: " + formatTime(totalTimeTaken) + ANSI_RESET);
        System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
    }

    // thReport 2023 II HRTA,ANJT,GOTO,KEJU,SBMA,PURA,KLBF,BAYU,BSML,INCO,MITI,ADMF
    // thReport 2023 III DOOH 15493
    // thReport 2023 Tahunan 15397 AALI,ACST,ADMF,ARNA,ASGR,BBCA,BBHI,BBNI,BFIN,BJTM,BMRI,CFIN,EAST,FUJI,HATM,INCO,ISAT,ITMG,LINK,LPPF,MCOR,MEGA,NIKL,PANS,PJAA,PNBN,UCID
    @Command(command = "thReport", description = "Read the financial data for the given year, period, and multiple kodeEmiten values.")
    public void thReport(String year, String period, Long usdIdrRate, String kodeEmitenList) {
        String[] kodeEmitens = kodeEmitenList.split(",");
        int currentCount = 0;
        int successfulCount = 0;
        int failedCount = 0;
        Map<String, String> failedDetails = new HashMap<>();
        long totalProcessingTime = 0;

        long startTimeOverall = System.currentTimeMillis(); // Start time for the entire process
        Map<String, TradingSummary> tradingSummary = tradingSummaryService.getTradingSummary();
        for (String kodeEmiten : kodeEmitens) {
            currentCount++;
            long startTime = System.currentTimeMillis(); // Start time for this kodeEmiten

            System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
            System.out.println(ANSI_CYAN + "🔄 Processing data for kodeEmiten: " + kodeEmiten.trim() +" ("+ currentCount + "/" + kodeEmitens.length + ") ..." + ANSI_RESET);
            try {
                excelReaderService.simplerReadExcel(year, period, kodeEmiten.trim(), usdIdrRate, tradingSummary);
                System.out.println(ANSI_GREEN + "✅ Successfully processed data for kodeEmiten: " + kodeEmiten.trim() + ANSI_RESET);
                successfulCount++;
            } catch (Exception e) {
                String errorMessage = e.getMessage();
                failedDetails.put(kodeEmiten.trim(), errorMessage);
                failedCount++;

                if (e instanceof IOException) {
                    System.out.println(ANSI_RED + "❌ Error processing data for kodeEmiten: " + kodeEmiten.trim() + ". Reason: " + errorMessage + ANSI_RESET);
                } else if (e instanceof InvalidFormatException) {
                    System.out.println(ANSI_RED + "⚠️ Invalid format encountered for kodeEmiten: " + kodeEmiten.trim() + ". Reason: " + errorMessage + ANSI_RESET);
                } else {
                    System.out.println(ANSI_RED + "❌ Unknown Error processing data for kodeEmiten: " + kodeEmiten.trim() + ". Reason: " + errorMessage + ANSI_RESET);
                }
            }
            long endTime = System.currentTimeMillis(); // End time for this kodeEmiten
            totalProcessingTime += (endTime - startTime);
            System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
        }

        // Summary
        System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
        System.out.println(ANSI_CYAN + "📊 SUMMARY 📊" + ANSI_RESET);
        System.out.println(ANSI_GREEN + "✅ Total Successful: " + successfulCount + ANSI_RESET);
        System.out.println(ANSI_RED + "❌ Total Failed: " + failedCount + ANSI_RESET);
        System.out.println(ANSI_BLUE + "📁 Total Files Processed: " + kodeEmitens.length + ANSI_RESET);

        if (failedCount > 0) {
            System.out.println(ANSI_RED + "\n❌ Failed Details:" + ANSI_RESET);
            for (Map.Entry<String, String> entry : failedDetails.entrySet()) {
                System.out.println(ANSI_RED + "❌ " + entry.getKey() + ": " + entry.getValue() + ANSI_RESET);
            }
        }

        long endTimeOverall = System.currentTimeMillis(); // End time for the entire process
        long totalTimeTaken = endTimeOverall - startTimeOverall;
        double averageTimePerKodeEmiten = (double) totalProcessingTime / kodeEmitens.length;

        System.out.println(ANSI_BLUE + "⏱️ Average Time Per KodeEmiten: " + formatTime((long) averageTimePerKodeEmiten) + ANSI_RESET);
        System.out.println(ANSI_BLUE + "⏱️ Total Time Taken: " + formatTime(totalTimeTaken) + ANSI_RESET);
        System.out.println(ANSI_YELLOW + "════════════════════════════════════════════════════════════════════════" + ANSI_RESET);
    }

    String formatTime(long milliseconds) {
        long totalSeconds = milliseconds / 1000;
        long minutes = totalSeconds / 60;
        long seconds = totalSeconds % 60;

        StringBuilder formattedTime = new StringBuilder();
        if (minutes > 0) {
            formattedTime.append(minutes).append("m ");
        }
        if (seconds > 0) {
            formattedTime.append(seconds).append("s");
        }
        return formattedTime.toString().trim(); // trim() to remove any trailing space
    }
}