package me.metallicgoat.datacombiner;

import java.awt.*;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextArea;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import javafx.application.Platform;
import javafx.concurrent.Task;


public class Main extends Application {

  private static final int INDEX_FILE_SAMPLES_ROW = 4;
  private static final int INDEX_FILE_DATES_ROW = 5;
  private static final int INDEX_FILE_DATA_SHEET = 2;

  // UNUSED
  private static final int INDEX_FILE_GRAPH_DATA_SHEET = 0;

  private static final String outputFileName = "updatedFile.xlsx";

  private File labDataFile;
  private File indexFile;
  private TextArea logArea;
  private Label fileLabel;
  private final HashMap<String, LocationTestData> labDataAnalTypesAndDate = new HashMap<>();
  private ProgressBar progressBar;
  private Button openOutputFileButton;


  public static void main(String[] args) {
    launch(args);
  }

  private void updateFileLabel() {
    String labFileName = labDataFile != null ? labDataFile.getName() : "No Lab Data file selected";
    String indexFileName = indexFile != null ? indexFile.getName() : "No Index file selected";
    fileLabel.setText("Lab Data: " + labFileName + " | Index: " + indexFileName);
  }

  @Override
  public void start(Stage stage) {
    stage.setTitle("Excel .xlsm Reader");

    Button selectLabDataFileButton = new Button("Select Lab Data File");
    Button selectIndexFileButton = new Button("Select Index File");
    Button submitButton = new Button("Submit");

    fileLabel = new Label("No file selected.");
    fileLabel.setStyle("-fx-text-fill: black;");

    logArea = new TextArea();
    logArea.setEditable(false);
    logArea.setWrapText(true);

    // NEW: Progress bar
    progressBar = new ProgressBar();
    progressBar.setVisible(false);
    progressBar.setPrefWidth(400);

    // NEW: Button to open output file
    openOutputFileButton = new Button("Open Output File");
    openOutputFileButton.setDisable(true);
    openOutputFileButton.setOnAction(e -> {
      try {
        Desktop.getDesktop().open(new File(outputFileName));
      } catch (IOException ex) {
        log("Unable to open file: " + ex.getMessage());
        showAlert("Could not open the file.");
      }
    });

    // Green border style for buttons (black text)
    String greenBorderStyle = "-fx-border-color: #4CAF50; -fx-border-width: 2px; -fx-text-fill: black; -fx-border-radius: 4px; -fx-background-radius: 4px;";

    selectLabDataFileButton.setStyle(greenBorderStyle);
    selectIndexFileButton.setStyle(greenBorderStyle);
    submitButton.setStyle(greenBorderStyle);
    openOutputFileButton.setStyle(greenBorderStyle);

    // Scrollable log area
    ScrollPane logScrollPane = new ScrollPane(logArea);
    logScrollPane.setFitToWidth(true);
    logScrollPane.setFitToHeight(true);
    logScrollPane.setPrefHeight(300);

    selectLabDataFileButton.setOnAction(e -> {
      labDataFile = promptSelectFile(stage);
      updateFileLabel();
    });

    selectIndexFileButton.setOnAction(e -> {
      indexFile = promptSelectFile(stage);
      updateFileLabel();
    });

    submitButton.setOnAction(e -> {
      if (labDataFile != null && indexFile != null) {
        progressBar.setVisible(true);
        progressBar.setProgress(-1); // Indeterminate bar

        Task<Void> task = new Task<>() {
          @Override
          protected Void call() throws Exception {
            try {
              log("Reading lab data file...");
              readLabDataFile(labDataFile);
              readArchiveFile(indexFile);
            } catch (Exception ex) {
              log("Error: " + ex.getMessage());
              Platform.runLater(() -> showAlert(ex.getMessage()));
            }
            return null;
          }

          @Override
          protected void succeeded() {
            progressBar.setVisible(false);
            openOutputFileButton.setDisable(false);
          }

          @Override
          protected void failed() {
            progressBar.setVisible(false);
            log("Task failed.");
          }
        };

        new Thread(task).start();
      } else {
        log("Both files must be selected.");
        showAlert("Please select both files first.");
      }
    });

    VBox layout = new VBox(10,
            selectLabDataFileButton,
            selectIndexFileButton,
            fileLabel,
            submitButton,
            progressBar,
            openOutputFileButton,
            logScrollPane
    );
    layout.setPadding(new Insets(20));
    layout.setStyle("-fx-font-family: 'Segoe UI', sans-serif;");
    VBox.setVgrow(logScrollPane, Priority.ALWAYS);

    stage.setScene(new Scene(layout, 600, 500));
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
      final Sheet dataSheet = workbook.getSheetAt(INDEX_FILE_DATA_SHEET);
      final Row sampleNameRow = dataSheet.getRow(INDEX_FILE_SAMPLES_ROW);
      final Row dateNameRow = dataSheet.getRow(INDEX_FILE_DATES_ROW);
      final List<String> seenTypes = new ArrayList<>();

      // Loop though all cells in the index file
      for (Cell cell : sampleNameRow) {
        if (cell.getCellType() == CellType.STRING && !seenTypes.contains(cell.getStringCellValue())) {
          seenTypes.add(cell.getStringCellValue());

          final String currId = tryToMatchID(cell.getStringCellValue());

          if (labDataAnalTypesAndDate.containsKey(currId)) {
            final LocationTestData locationTestData = labDataAnalTypesAndDate.get(currId);
            final Date newestDate = locationTestData.dataDate;

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
                  if (date.getYear() == newestDate.getYear() && date.getMonth() == newestDate.getMonth()) {
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
            } else if (needsUpdating) {
              log(currId + " - NEEDS TO BE UPDATED");

              createNewCol(dataSheet, currColIndex);

              final Cell newDateCell = dateNameRow.getCell(currColIndex);

              newDateCell.setCellValue(newestDate);

              for (Row row : dataSheet) {
                final Cell paramCell = row.getCell(0);

                if (paramCell != null && paramCell.getCellType() == CellType.STRING){
                  final String paramId = paramCell.getStringCellValue();

                  final String data = locationTestData.getDataByParam(paramId);

                  // Write data to index file
                  writeToCell(row.getCell(currColIndex), data);
                }
              }

              labDataAnalTypesAndDate.remove(currId);

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

  private String tryToMatchID(String possibleId) {
    possibleId = possibleId.replace(" ", "").toLowerCase();

    for (String id : labDataAnalTypesAndDate.keySet()) {
      if (possibleId.contains(id.replace(" ", "").toLowerCase())) {
        return id;
      }
    }

    return possibleId;
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

        // hit the end maybe?
        if (sampleId.isEmpty()) {
          nullCount++;
          continue;
        }

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

          // Some cells are empty because there are multiple tests and only one is used
          if (sheet.getRow(currRow) == null || sheet.getRow(currRow).getCell(0) == null) {
            nullRowCount++;
            continue;
          }

          nullRowCount = 0;

          final String paramName = getCellValue(sheet.getRow(currRow).getCell(0));

          String paramValue = getCellValue(sheet.getRow(currRow).getCell(currColumn));

          int it = 0;
          while (paramValue.isEmpty()) {
            it += 1;

            final Cell nextCell = sheet.getRow(currRow + it).getCell(currColumn);

            paramValue = nextCell == null ? "" : getCellValue(nextCell);
          }

          locationTestData.addData(paramName, paramValue);
        }
      }

    } catch (Exception ex) {
      log("Error: " + ex.getMessage());
      ex.printStackTrace();
      throw new Exception("Lab data seems fucked up. Check the cell constants and format.");
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


  private void writeToCell(Cell cell, String value) {
    if (value == null || value.trim().isEmpty()) {
      cell.setBlank();
      return;
    }

    try {
      // fix some things copied as strings
      double numericValue = Double.parseDouble(value);
      cell.setCellValue(numericValue);
    } catch (NumberFormatException e) {
      cell.setCellValue(value);
    }
  }

  private void createNewCol(Sheet sheet, int colIndex) {
    Workbook workbook = sheet.getWorkbook();

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


  private void showAlert(String msg) {
    Alert alert = new Alert(Alert.AlertType.WARNING, msg, ButtonType.OK);
    alert.showAndWait();
  }
}
