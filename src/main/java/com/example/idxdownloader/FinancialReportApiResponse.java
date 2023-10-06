package com.example.idxdownloader;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

import java.util.List;
import java.util.Map;

@Data
public class FinancialReportApiResponse {
    @JsonProperty("Search")
    private Map<String, String> search;
    @JsonProperty("ResultCount")
    private int resultCount;
    @JsonProperty("Results")
    private List<Result> results;
}