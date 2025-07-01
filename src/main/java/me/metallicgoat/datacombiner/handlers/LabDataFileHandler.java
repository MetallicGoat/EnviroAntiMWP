package me.metallicgoat.datacombiner.handlers;

import me.metallicgoat.datacombiner.util.Helpers;
import me.metallicgoat.datacombiner.util.LocationTestData;
import me.metallicgoat.datacombiner.util.Util;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

public class LabDataFileHandler {

    // Reads the lab data file and populates the labTestDataByLocationId map
    public static void readAndStoreData(File file) throws Exception {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            final Sheet sheet = workbook.getSheetAt(0);

            // Find the wells
            // Get sample IDs Start at E 11
            Row sampleRow = sheet.getRow(10);
            int sampleColumn = 4;
            final Cell sampleIdCell = sampleRow.getCell(sampleColumn);

            if (sampleIdCell == null || sampleIdCell.getCellType() == CellType.STRING && !sampleIdCell.getStringCellValue().trim().equals("Sample ID")) {
                throw new Exception("Sample ID not found in expected cell. Did you select the correct lab data file?");
            }

            int currColumn = sampleColumn;
            int nullCount = 0;

            // Sometimes there are gaps for multiple tests
            // If there is more than 10, it's probably the end of the data
            while (nullCount < 10) {
                currColumn += 1;

                if (sampleRow.getCell(currColumn) == null) {
                    nullCount++;
                    continue;
                }

                final String sampleId = sheet.getRow(sampleRow.getRowNum()).getCell(currColumn).getStringCellValue();

                // hit the end maybe?
                if (sampleId.isEmpty()) {
                    nullCount++;
                    continue;
                }

                final String sampleIdDate = sheet.getRow(sampleRow.getRowNum() + 1).getCell(currColumn).getStringCellValue();
                final Date date = Date.from(LocalDate.parse(sampleIdDate).atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());

                final LocationTestData locationTestData = new LocationTestData();
                Util.storeLabData(sampleId.toUpperCase(), locationTestData);
                locationTestData.dataDate = date;

                nullCount = 0;

                int currRow = sampleRow.getRowNum() + 2;
                int nullRowCount = 0;

                while (nullRowCount < 10) {
                    currRow += 1;

                    // Some cells are empty because there are multiple tests and only one is used
                    if (sheet.getRow(currRow) == null || sheet.getRow(currRow).getCell(0) == null) {
                        nullRowCount++;
                        continue;
                    }

                    nullRowCount = 0;

                    final String paramName = Helpers.getCellValue(sheet.getRow(currRow).getCell(0));

                    String paramValue = Helpers.getCellValue(sheet.getRow(currRow).getCell(currColumn));

                    int it = 0;

                    // If the value is empty, we need to keep looking down the column
                    // This is for cases where data is not aligned due to different tests types
                    while (paramValue.isEmpty()) {
                        it += 1;
                        final Cell nextCell = sheet.getRow(currRow + it).getCell(currColumn);
                        paramValue = nextCell == null ? "" : Helpers.getCellValue(nextCell);
                    }

                    locationTestData.addData(paramName, paramValue);
                }
            }

        } catch (Exception ex) {
            Util.getUI().log("Error: " + ex.getMessage());
            throw new Exception("The lab data file is not in the expected format. ");
        }
    }
}
