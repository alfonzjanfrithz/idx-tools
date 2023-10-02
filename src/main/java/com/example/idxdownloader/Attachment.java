package com.example.idxdownloader;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

@Data
public class Attachment {
    @JsonProperty("Emiten_Code")
    private String emitenCode;

    @JsonProperty("File_ID")
    private String fileId;

    @JsonProperty("File_Modified")
    private String fileModified;

    @JsonProperty("File_Name")
    private String fileName;

    @JsonProperty("File_Path")
    private String filePath;

    @JsonProperty("File_Size")
    private long fileSize;

    @JsonProperty("File_Type")
    private String fileType;

    @JsonProperty("Report_Period")
    private String reportPeriod;

    @JsonProperty("Report_Type")
    private String reportType;

    @JsonProperty("Report_Year")
    private String reportYear;

    @JsonProperty("NamaEmiten")
    private String namaEmiten;
}