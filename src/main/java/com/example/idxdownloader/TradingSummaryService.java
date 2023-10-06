package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.springframework.stereotype.Service;

import java.util.Map;
import java.util.stream.Collectors;

@AllArgsConstructor
@Service
public class TradingSummaryService {
    private FinancialStatementService financialStatementService;

    public Map<String, TradingSummary> getTradingSummary() {
        TradingSummaryApiResponse tradingSummaryApiResponse = financialStatementService.fetchTradingSummary();
        return tradingSummaryApiResponse.getData().stream()
                .collect(Collectors.toMap(TradingSummary::getStockCode, tradingSummary -> tradingSummary));
    }
}
