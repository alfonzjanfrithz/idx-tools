package com.example.idxdownloader;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.springframework.stereotype.Service;
import org.springframework.web.util.UriComponentsBuilder;

@Service
public class FinancialStatementService {
    public FinancialReportApiResponse fetchFinancialReport(int year, String periode, String kodeEmiten) {
        System.setProperty("webdriver.chrome.driver", ".\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        String url = UriComponentsBuilder
                .fromHttpUrl("https://idx.co.id/primary/ListedCompany/GetFinancialReport")
                .queryParam("indexFrom", 1)
                .queryParam("pageSize", 1)
                .queryParam("year", year)
                .queryParam("reportType", "rdf")
                .queryParam("EmitenType", "s")
                .queryParam("periode", periode)
                .queryParam("kodeEmiten", kodeEmiten)
                .queryParam("SortColumn", "KodeEmiten")
                .queryParam("SortOrder", "asc")
                .toUriString();

        driver.get(url);
        String jsonContent = driver.findElement(By.tagName("pre")).getText();
        driver.quit();

        ObjectMapper objectMapper = new ObjectMapper();
        try {
            return objectMapper.readValue(jsonContent, FinancialReportApiResponse.class);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
            return null;
        }
    }

    public TradingSummaryApiResponse fetchTradingSummary() {
        System.setProperty("webdriver.chrome.driver", ".\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        String url = UriComponentsBuilder
                .fromHttpUrl("https://idx.co.id/primary/TradingSummary/GetStockSummary")
                .queryParam("length", 9999)
                .queryParam("start", 0)
                .toUriString();

        driver.get(url);
        String jsonContent = driver.findElement(By.tagName("pre")).getText();
        driver.quit();

        ObjectMapper objectMapper = new ObjectMapper();
        try {
            return objectMapper.readValue(jsonContent, TradingSummaryApiResponse.class);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
            return null;
        }
    }
}
