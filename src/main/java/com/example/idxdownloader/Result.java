package com.example.idxdownloader;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

import java.util.List;

@Data
public class Result {

    @JsonProperty("KodeEmiten")
    private String kodeEmiten;

    @JsonProperty("File_Modified")
    private String fileModified;

    @JsonProperty("Report_Period")
    private String reportPeriod;

    @JsonProperty("Report_Year")
    private String reportYear;

    @JsonProperty("NamaEmiten")
    private String namaEmiten;

    @JsonProperty("Attachments")
    private List<Attachment> attachments;
}