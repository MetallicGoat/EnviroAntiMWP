package me.metallicgoat.datacombiner.util;

import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Helpers {

    // Has file picker support for linux
    public static void openFile(File file) throws Exception {
        try {
            if (!file.exists()) {
                throw new Exception("File doesn't exist.");
            }

            if (System.getProperty("os.name").toLowerCase().contains("linux")) {
                new ProcessBuilder("xdg-open", file.getAbsolutePath()).start();
            } else {
                Desktop.getDesktop().open(file);
            }

        } catch (Exception ex) {
            throw new Exception("Failed to open file: " + ex.getMessage(), ex);
        }
    }

    public static String columnIndexToExcelLetter(int column) {
        final StringBuilder letter = new StringBuilder();

        while (column >= 0) {
            int remainder = column % 26;
            letter.insert(0, (char) (remainder + 'A'));
            column = (column / 26) - 1;
        }

        return letter.toString();
    }

    // Index Files Sometimes have weird pram names with extra additions
    public static String normalizeSampleId(String input) {
        if (input == null) return "";
        return input
                .toLowerCase()
                .replaceAll("\\b(cont'd|continued|duplicate)\\b", "")
                .replaceAll("[^a-z0-9]", "")
                .trim();
    }

    public static String getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case BLANK -> "";
            default -> "UNKNOWN";
        };
    }

    // writes as a string or numeric value
    public static void writeToCell(Cell cell, String value) {
        if (value == null || value.trim().isEmpty()) {
            cell.setBlank();
            return;
        }

        Workbook workbook = cell.getSheet().getWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        CellStyle dateStyle = workbook.createCellStyle();
        short dateFormat = creationHelper.createDataFormat().getFormat("MMM-yy");
        dateStyle.setDataFormat(dateFormat);

        try {
            // Try parsing as a date first
            final Date parsedDate = new SimpleDateFormat("MMM-yy").parse(value);

            cell.setCellValue(parsedDate);
            cell.setCellStyle(dateStyle);

        } catch (Exception ignored) {
            try {
                // Fall back to numeric
                double numericValue = Double.parseDouble(value);
                cell.setCellValue(numericValue);
            } catch (NumberFormatException e) {
                // Otherwise just write as a string
                cell.setCellValue(value);
            }
        }
    }


    // Creates a new column, coping the style from the left cells
    // Try to be consistent with the index file
    public static void createNewColWithCopiedStyles(Sheet sheet, int colIndex) {
        final Workbook workbook = sheet.getWorkbook();

        for (Row row : sheet) {
            final Cell leftCell = row.getCell(colIndex - 1);
            final Cell newCell = row.createCell(colIndex);

            if (leftCell != null && leftCell.getCellStyle() != null) {
                final CellStyle newStyle = workbook.createCellStyle();
                newStyle.cloneStyleFrom(leftCell.getCellStyle());
                newCell.setCellStyle(newStyle);
            }
        }
    }
}
