package com.example.idxdownloader;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

@Service
public class FileDownloadService {

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