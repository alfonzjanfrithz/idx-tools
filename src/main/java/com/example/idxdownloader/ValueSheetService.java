package com.example.idxdownloader;

import lombok.AllArgsConstructor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.HashMap;

@Service
@AllArgsConstructor
public class ValueSheetService {
    private ExcelReaderService excelReaderService;
    private static final String BASE_DIRECTORY = System.getProperty("user.home") + "\\.idx-tmp\\value-sheet\\";

    public String getFilePath(String kodeEmiten) {
        return BASE_DIRECTORY + kodeEmiten + "_VR_Sheet_v5_0.xlsx";
    }
    public void populateTemplate(String kodeEmiten) throws IOException, InvalidFormatException {
        FinancialData year2022 = excelReaderService.getFinancialData("2022", "Tahunan", kodeEmiten);

        // Path to the template file
        String templatePath = "VR_Sheet_v5_0.xlsx"; // Update the path accordingly

        // Open the template workbook
        FileInputStream fis = new FileInputStream(new File(templatePath));
        Workbook workbook = new XSSFWorkbook(fis);

        // Get the sheet named "Ratios"
        Sheet sheet = workbook.getSheet("Ratios");
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet 'Ratios' not found in the template.");
        }

        writeCell(sheet, "B2", kodeEmiten + " Financial Data");
        writeCell(sheet, "C6", year2022.getRevenue());
        // Save the populated data to a new file
        String newFileName = getFilePath(kodeEmiten);
        FileOutputStream fos = new FileOutputStream(newFileName);
        workbook.write(fos);

        // Close resources
        fos.close();
        workbook.close();
        fis.close();
    }

    public void writeCell(Sheet sheet, String cellReference, Object value) {
        int rowIndex = Integer.parseInt(cellReference.replaceAll("[^0-9]", "")) - 1; // Extract row number and convert to 0-based index
        int colIndex = columnToIndex(cellReference.replaceAll("[0-9]", "")); // Extract column letter and convert to index

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

        if (value instanceof Double) {
            cell.setCellValue((Double) value);
        }else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        }
    }

    private int columnToIndex(String column) {
        int index = 0;
        char[] chars = column.toUpperCase().toCharArray();
        for (int i = 0; i < chars.length; i++) {
            index *= 26;
            index += chars[i] - 'A' + 1;
        }
        return index - 1; // Convert to 0-based index
    }
}
