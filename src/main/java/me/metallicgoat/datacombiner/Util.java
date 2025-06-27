package me.metallicgoat.datacombiner;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;

public class Util {

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
        StringBuilder letter = new StringBuilder();
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
}
