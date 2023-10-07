package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class FinancialData {
    private boolean IDRCurrency;
    private long multiplier;
    private Double totalLiabilities;
    private Double totalEquities;
    private Double netProfit;
    private Double netProfitLastYear;

    private Double revenue;
    private Double cogs;
}
