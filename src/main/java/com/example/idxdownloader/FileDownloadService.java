package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

@Service
@AllArgsConstructor
public class FileDownloadService {
    private static final String BASE_DIRECTORY = System.getProperty("user.home") + "\\.idx-tmp\\";

    private FinancialStatementService financialStatementService;

    public String getFilePath(String year, String period, String kodeEmiten) {
        return BASE_DIRECTORY + "FinancialStatement-" + year + "-" + period + "-" + kodeEmiten + ".xlsx";
    }

    public boolean fileExists(String filePath) {
        File file = new File(filePath);
        return file.exists();
    }

    public void downloadFS(int year, String periode, String kodeEmiten) {
        if ("I".equalsIgnoreCase(periode)) {
            periode = "tw1";
        } else if ("II".equalsIgnoreCase(periode)) {
            periode = "tw2";
        } else if ("III".equalsIgnoreCase(periode)) {
            periode = "tw3";
        } else {
            periode = "audit";
        }
        ApiResponse apiResponse = financialStatementService.fetchData(year, periode, kodeEmiten);
        List<Attachment> attachmentExcel = filterAttachmentsByFileType(apiResponse, "xlsx");
        Optional<String> link = attachmentExcel.stream().map(attachment -> "https://idx.co.id" + attachment.getFilePath()).findAny();
        downloadFile(link.get());
    }

    public List<Attachment> filterAttachmentsByFileType(ApiResponse apiResponse, String fileType) {
        return apiResponse.getResults().stream()
                .flatMap(result -> result.getAttachments().stream())
                .filter(attachment -> (attachment.getFileType().contains(fileType)))
                .collect(Collectors.toList());
    }

    public void downloadFile(String downloadLink) {
        ChromeOptions options = new ChromeOptions();

        // Set default download directory
        Map<String, Object> chromePrefs = new HashMap<>();
        chromePrefs.put("download.default_directory", System.getProperty("user.home") + "\\.idx-tmp");
        chromePrefs.put("download.prompt_for_download", false);  // Avoid multiple download dialogs
        options.setExperimentalOption("prefs", chromePrefs);

        // Disable pop-ups blocking to allow for download
        options.addArguments("--disable-popup-blocking");

        WebDriver driver = new ChromeDriver(options);

        // Navigate to download link
        driver.get(downloadLink);

        // Wait for the download to complete
        String fileName = extractFileName(downloadLink);
        waitForDownloadToComplete(fileName);

        // Close the driver
        driver.quit();
    }

    private String extractFileName(String downloadLink) {
        return downloadLink.substring(downloadLink.lastIndexOf('/') + 1);
    }

    private void waitForDownloadToComplete(String fileName) {
        File tempFile;
        do {
            try {
                Thread.sleep(1000);  // wait for a second before checking again
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
            tempFile = new File(System.getProperty("user.home") + "\\idx-tmp\\" + fileName + ".crdownload");
        } while (tempFile.exists());
    }
}