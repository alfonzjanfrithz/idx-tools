package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellStyle;

@AllArgsConstructor
@Builder
@Getter
@Setter
public class WorkbookStyle {
    CellStyle currencyStyle;
    CellStyle percentageStyle;
    CellStyle decimalStyle;
    CellStyle dateStyle;
}
