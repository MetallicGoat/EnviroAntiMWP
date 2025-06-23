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

    public static Workbook createNewWorkBook() {
        try (Workbook workbook = new XSSFWorkbook(); Workbook workbook2 = new XSSFWorkbook()) {
            return workbook;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Failed to create new workbook", e);
        }
    }
}
