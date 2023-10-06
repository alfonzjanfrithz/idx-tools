package com.example.idxdownloader;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

import java.util.List;

@Data
public class TradingSummaryApiResponse {

    @JsonProperty("draw")
    private int draw;

    @JsonProperty("recordsTotal")
    private int recordsTotal;

    @JsonProperty("recordsFiltered")
    private int recordsFiltered;

    @JsonProperty("data")
    private List<TradingSummary> data;



}