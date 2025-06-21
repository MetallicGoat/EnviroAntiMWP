package me.metallicgoat.datacombiner;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main extends Application {

  private static final int LAB_DATA_SAMPLES_ROW = 2;
  private static final int LAB_DATA_DATES_ROW = 3;
  private static final int INDEX_FILE_DATA_SHEET = 1;
  private static final int INDEX_FILE_GRAPH_DATA_SHEET = 0;

  private static final String outputFileName = "updatedFile.xlsx";

  private File labDataFile;
  private File indexFile;
  private TextArea logArea;
  private Label fileLabel;

  private final HashMap<String, LocationTestData> labDataAnalTypesAndDate = new HashMap<>();
  private int commonTypes = 0;

  public static void main(String[] args) {
    launch(args);
  }

  @Override
  public void start(Stage stage) {
    stage.setTitle("Excel .xlsm Reader");

    Button selectLabDataFileButton = new Button("Select Lab Data File");
    Button selectIndexFileButton = new Button("Select Index File");
    Button submitButton = new Button("Submit");

    fileLabel = new Label("No file selected.");
    logArea = new TextArea();
    logArea.setEditable(false);
    logArea.setWrapText(true);

    ScrollPane logScrollPane = new ScrollPane(logArea);
    logScrollPane.setFitToWidth(true);
    logScrollPane.setFitToHeight(true);
    logScrollPane.setPrefHeight(300);

    selectLabDataFileButton.setOnAction(e -> {
      labDataFile = promptSelectFile(stage);
    });
    selectIndexFileButton.setOnAction(e -> {
      indexFile = promptSelectFile(stage);
    });

    submitButton.setOnAction(e -> {
      if (labDataFile != null) {
        log("Reading lab data file...");
        try {
          // Load the Data file into memory
          readLabDataFile(labDataFile);

          // Compare Data to Lab file
          readArchiveFile(indexFile);

        } catch (Exception ex) {
          log("Error reading lab data file: " + ex.getMessage());
          showAlert(ex.getMessage());
        } catch (Throwable ex) {
          showAlert(ex.getMessage());        }

      } else {
        log("No file selected.");
        showAlert("Please select a file first.");
      }
    });

    VBox layout = new VBox(10,
        selectLabDataFileButton,
        selectIndexFileButton,
        fileLabel,
        submitButton,
        logScrollPane
    );
    layout.setPadding(new Insets(20));
    layout.setStyle("-fx-font-family: 'Segoe UI', sans-serif;");

    stage.setScene(new Scene(layout, 600, 450));
    stage.show();
  }

  // File Picker UI
  private File promptSelectFile(Stage stage) {
    FileChooser fileChooser = new FileChooser();
    fileChooser.setTitle("Select Excel Macro File");

    File selectedFile = fileChooser.showOpenDialog(stage);

    if (selectedFile != null) {
      fileLabel.setText("Selected: " + selectedFile.getName());
      log("File selected: " + selectedFile.getAbsolutePath());
    }

    return selectedFile;
  }

  private void readArchiveFile(File file) throws Exception {
    try (FileInputStream fis = new FileInputStream(file);
         Workbook workbook = new XSSFWorkbook(fis)) {

      final Calendar calendar = new GregorianCalendar();
      final Sheet graphSheet = workbook.getSheetAt(INDEX_FILE_GRAPH_DATA_SHEET);
      final Sheet dataSheet = workbook.getSheetAt(INDEX_FILE_DATA_SHEET);


      final Row sampleNameRow = dataSheet.getRow(LAB_DATA_SAMPLES_ROW);
      final Row dateNameRow = dataSheet.getRow(LAB_DATA_DATES_ROW);

      final List<String> seenTypes = new ArrayList<>();

      // Loop though all cells in the index file
      for (Cell cell : sampleNameRow) {
        if (cell.getCellType() == CellType.STRING && !seenTypes.contains(cell.getStringCellValue())) {
          seenTypes.add(cell.getStringCellValue());

          final String currId = cell.getStringCellValue().toUpperCase();

          if (labDataAnalTypesAndDate.containsKey(currId)) {
            final LocationTestData locationTestData = labDataAnalTypesAndDate.get(currId);
            final Date newestData = locationTestData.dataDate;

            commonTypes++;

            // Go down one cell in the col, and loop until we find the latest date
            int currColIndex = cell.getColumnIndex();
            int curYear = 0;
            boolean keepLooking = true;
            boolean needsUpdating = true;
            String issue = null;

            while (keepLooking) {
              final Cell dateCell = dateNameRow.getCell(currColIndex);

              try {
                final Date date = dateCell.getDateCellValue();
                calendar.setTime(date);
                final int year = calendar.get(Calendar.YEAR);

                // Once the date is smaller, it is the next sample, gone to far
                if (curYear <= year) {
                  if (date.getYear() == newestData.getYear() && date.getMonth() == newestData.getMonth()) {
                    needsUpdating = false;
                    break;
                  }

                  curYear = year;
                  currColIndex++;
                } else {
                  keepLooking = false;
                }

              } catch (Exception ex) {
                keepLooking = false;
                issue = "FAILED TO SEARCH THOUGH DATES - ARE THEY IN DATE FORMAT?";
              }
            }

            if (issue != null) {
              log(currId + " - HAS AN ISSUE - " + issue);
              continue;
            }


            if (needsUpdating) {
              log(currId + " - NEEDS TO BE UPDATED");

              createNewCol(dataSheet, currColIndex);

              for (Row row : dataSheet) {
                final Cell paramCell = row.getCell(0);

                if (paramCell != null && paramCell.getCellType() == CellType.STRING){
                  final String paramId = paramCell.getStringCellValue();

                  final String data = locationTestData.getDataByParam(paramId);

                  // Write data to index file
                  row.getCell(currColIndex).setCellValue(data);

                  System.out.println(paramId + " - " + data);
                }
              }

            } else {
              log(currId + " - UP TO DATE ALREADY");
            }
          }
        }
      }

      try (FileOutputStream fos = new FileOutputStream(outputFileName)) {
        workbook.write(fos);
        log("Updated file written to: " + outputFileName);
      }

    } catch (Exception ex) {
      log("Error: " + ex.getMessage());
      ex.printStackTrace();
      throw new Exception("Failed to load the Index File, check cell constants ");
    }
  }

  private void readLabDataFile(File file) throws Exception {
    try (FileInputStream fis = new FileInputStream(file);
         Workbook workbook = new XSSFWorkbook(fis)) {

      final Sheet sheet = workbook.getSheetAt(0);

      // Find the wells
      // Get sample IDs Start at E 11
      Row sampleRow = sheet.getRow(10);
      int sampleColumn = 4;
      final Cell sampleIdCell = sampleRow.getCell(sampleColumn);

      if (sampleIdCell == null || sampleIdCell.getCellType() == CellType.STRING && !sampleIdCell.getStringCellValue().trim().equals("Sample ID")) {
        throw new Exception("ME IS SO CONFUSED? Sample ID not found in expected cell.");
      }

      int currColumn = sampleColumn;
      int nullCount = 0;

      while (nullCount < 10) {
        currColumn += 1;

        if (sampleRow.getCell(currColumn) == null) {
          nullCount++;
          continue;
        }

        final String sampleId = sheet.getRow(sampleRow.getRowNum()).getCell(currColumn).getStringCellValue();
        final String sampleIdDate = sheet.getRow(sampleRow.getRowNum() + 1).getCell(currColumn).getStringCellValue();
        final Date date = Date.from(LocalDate.parse(sampleIdDate).atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());


        final LocationTestData locationTestData = new LocationTestData();
        labDataAnalTypesAndDate.put(sampleId.toUpperCase(), locationTestData);
        locationTestData.dataDate = date;

        nullCount = 0;

        int currRow = sampleRow.getRowNum() + 2;
        int nullRowCount = 0;

        while (nullRowCount < 10) {
          currRow += 1;

          if (sheet.getRow(currRow) == null || sheet.getRow(currRow).getCell(0) == null) {
            nullRowCount++;
            System.out.println("Null row at: " + currRow);
            continue;
          }

          nullRowCount = 0;

          final String paramName = getCellValue(sheet.getRow(currRow).getCell(0));

          String paramValue = getCellValue(sheet.getRow(currRow).getCell(currColumn));

          int it = 0;
          while (paramValue.isEmpty()) {
            it += 1;

            Cell nextCell = sheet.getRow(currRow + it).getCell(currColumn);

            paramValue = nextCell == null ? "" : getCellValue(nextCell);
          }

          System.out.println(sampleId + ": " + paramName + " - " + paramValue);
          locationTestData.addData(paramName, paramValue);

        }
      }

    } catch (Exception ex) {
      log("Error: " + ex.getMessage());
      ex.printStackTrace();
      throw new Exception("Lab data seems fucked up ");
    }
  }

  private String getCellValue(Cell cell) {
    return switch (cell.getCellType()) {
      case STRING -> cell.getStringCellValue().trim();
      case NUMERIC -> String.valueOf(cell.getNumericCellValue());
      case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
      case FORMULA -> cell.getCellFormula();
      case BLANK -> "";
      default -> "UNKNOWN";
    };
  }

  private void log(String message) {
    logArea.appendText(message + "\n");
  }

  private void createNewCol(Sheet sheet, int colIndex) {
    for (Row row : sheet) {
      Cell cell = row.createCell(colIndex);
      if (row.getRowNum() == 0) {
        // cell.setCellValue("New Column"); // Header row
      } else {
        //cell.setCellValue("Value " + row.getRowNum()); // Example data
      }
    }
  }

  private void showAlert(String msg) {
    Alert alert = new Alert(Alert.AlertType.WARNING, msg, ButtonType.OK);
    alert.showAndWait();
  }
}
