package com.example.idxdownloader;

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

@Service
public class ValueSheetService {
    private static final String BASE_DIRECTORY = System.getProperty("user.home") + "\\.idx-tmp\\value-sheet\\";

    public String getFilePath(String kodeEmiten) {
        return BASE_DIRECTORY + kodeEmiten + "_VR_Sheet_v5_0.xlsx";
    }
    public void populateTemplate(String kodeEmiten) throws IOException {
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

        // Populate cell B2
        Row row = sheet.getRow(1); // 0-based index, so 1 is the second row
        if (row == null) {
            row = sheet.createRow(1);
        }
        Cell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // 0-based index, so 1 is column B
        cell.setCellValue(kodeEmiten + " Financial Data");

        // Save the populated data to a new file
        String newFileName = getFilePath(kodeEmiten);
        FileOutputStream fos = new FileOutputStream(newFileName);
        workbook.write(fos);

        // Close resources
        fos.close();
        workbook.close();
        fis.close();
    }
}
