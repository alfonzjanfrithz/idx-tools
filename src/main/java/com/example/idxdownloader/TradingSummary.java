package com.example.idxdownloader;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

@Data
public class TradingSummary {

    @JsonProperty("No")
    private Long no;

    @JsonProperty("IDStockSummary")
    private Long idStockSummary;

    @JsonProperty("Date")
    private String date;

    @JsonProperty("StockCode")
    private String stockCode;

    @JsonProperty("StockName")
    private String stockName;

    @JsonProperty("Remarks")
    private String remarks;

    @JsonProperty("Previous")
    private Long previous;

    @JsonProperty("OpenPrice")
    private Long openPrice;

    @JsonProperty("FirstTrade")
    private Long firstTrade;

    @JsonProperty("High")
    private Long high;

    @JsonProperty("Low")
    private Long low;

    @JsonProperty("Close")
    private Long close;

    @JsonProperty("Change")
    private Long change;

    @JsonProperty("Volume")
    private Long volume;

    @JsonProperty("Value")
    private Long value;

    @JsonProperty("Frequency")
    private Long frequency;

    @JsonProperty("IndexIndividual")
    private double indexIndividual;

    @JsonProperty("Offer")
    private Long offer;

    @JsonProperty("OfferVolume")
    private Long offerVolume;

    @JsonProperty("Bid")
    private Long bid;

    @JsonProperty("BidVolume")
    private Long bidVolume;

    @JsonProperty("ListedShares")
    private Long listedShares;

    @JsonProperty("TradebleShares")
    private Long tradebleShares;

    @JsonProperty("WeightForIndex")
    private Long weightForIndex;

    @JsonProperty("ForeignSell")
    private Long foreignSell;

    @JsonProperty("ForeignBuy")
    private Long foreignBuy;

    @JsonProperty("DelistingDate")
    private String delistingDate;

    @JsonProperty("NonRegularVolume")
    private Long nonRegularVolume;

    @JsonProperty("NonRegularValue")
    private Long nonRegularValue;

    @JsonProperty("NonRegularFrequency")
    private Long nonRegularFrequency;

    @JsonProperty("persen")
    private Double persen;  // This can be changed to a specific type if known

    @JsonProperty("percentage")
    private Double percentage;  // This can be changed to a specific type if known

}
