package me.metallicgoat.datacombiner;

import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javafx.geometry.Pos;

import java.io.*;


public class Main extends Application {

    private static final String APP_NAME = "EnvioAntiMWP";

    private static final int INDEX_FILE_SAMPLES_ROW = 1;
    private static final int INDEX_FILE_DATES_ROW = 2;
    private static final int INDEX_FILE_DATA_SHEET = 1;

    private static final String editedIndexFileName = "output/UpdatedFile.xlsx";
    private static final String changesOnlyFileName = "output/ChangesOnly.xlsx";

    private File labDataFile;
    private File indexFile;
    private TextArea logArea;
    private Button submitButton;
    private final HashMap<String, LocationTestData> labTestDataByLocationId = new HashMap<>();
    private Button openFileButton;
    private Button openChangesOnlyFileButton;


    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage stage) {
        stage.setTitle("Index Updater");

        final Label labDataLabel = new Label("Lab Data File:");
        final TextField labDataField = new TextField();
        labDataField.setPromptText("No file selected");
        labDataField.setEditable(false);
        // Add rounding for TextFields
        labDataField.setStyle("-fx-background-radius: 8;");

        final Button browseLabButton = new Button("Browse");
        browseLabButton.setStyle("-fx-background-radius: 8;");

        final Label indexFileLabel = new Label("Index File:");
        final TextField indexFileField = new TextField();
        indexFileField.setPromptText("No file selected");
        indexFileField.setEditable(false);
        indexFileField.setStyle("-fx-background-radius: 8;");

        final Button browseIndexButton = new Button("Browse");
        browseIndexButton.setStyle("-fx-background-radius: 8;");



        final GridPane fileGrid = new GridPane();
        fileGrid.setHgap(10);
        fileGrid.setVgap(10);
        final ColumnConstraints col1 = new ColumnConstraints();
        col1.setMinWidth(100);
        final ColumnConstraints col2 = new ColumnConstraints();
        col2.setHgrow(Priority.ALWAYS);  // Allow the text fields to expand
        final ColumnConstraints col3 = new ColumnConstraints();
        col3.setMinWidth(100);
        fileGrid.getColumnConstraints().addAll(col1, col2, col3);

        labDataField.setMaxWidth(Double.MAX_VALUE);
        indexFileField.setMaxWidth(Double.MAX_VALUE);

        fileGrid.add(labDataLabel, 0, 0);
        fileGrid.add(labDataField, 1, 0);
        fileGrid.add(browseLabButton, 2, 0);
        fileGrid.add(indexFileLabel, 0, 1);
        fileGrid.add(indexFileField, 1, 1);
        fileGrid.add(browseIndexButton, 2, 1);

        final TitledPane fileSection = new TitledPane("Select Files", fileGrid);
        fileSection.setCollapsible(false);
        fileSection.setStyle("-fx-background-radius: 8;");

        submitButton = new Button("Process and Update Index File");
        submitButton.setMaxWidth(Double.MAX_VALUE);
        submitButton.setStyle("-fx-font-size: 14px; -fx-background-color: #038262; -fx-text-fill: white; -fx-background-radius: 8;");
        submitButton.setOnMouseEntered(e ->
                submitButton.setStyle("-fx-font-size: 14px; -fx-background-color: #0a8c6c; -fx-text-fill: white; -fx-background-radius: 8;")
        );
        submitButton.setOnMouseExited(e ->
                submitButton.setStyle("-fx-font-size: 14px; -fx-background-color: #038262; -fx-text-fill: white; -fx-background-radius: 8;")
        );

        logArea = new TextArea();
        logArea.setEditable(false);
        logArea.setWrapText(true);
        logArea.setStyle("-fx-font-family: 'Consolas'; -fx-background-radius: 8;");
        logArea.setPrefHeight(300);

        Label creditLabel = new Label("Created by Christian Azzam and Braydon Affleck · Written by Christian Azzam.");
        creditLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #777;");

        final TitledPane logPane = new TitledPane("Log Output", logArea);
        logPane.setCollapsible(false);
        logPane.setStyle("-fx-background-radius: 8;");

        // Open Updated File button placed below log output
        final Button openFileButton = new Button("Open Updated File");
        openFileButton.setDisable(true);
        openFileButton.setMaxWidth(Double.MAX_VALUE);
        openFileButton.setStyle("-fx-background-color: #cccccc; -fx-text-fill: black; -fx-font-size: 14px; -fx-background-radius: 8;");

        openFileButton.setOnAction(e -> {
            try {
                Util.openFile(new File(editedIndexFileName));
            } catch (Exception ex) {
                showAlert(ex.getMessage(), Alert.AlertType.ERROR);
                log(ex.getMessage());
            }
        });

        Label orLabel = new Label("or");

        openChangesOnlyFileButton = new Button("Open Change File");
        openChangesOnlyFileButton.setDisable(true);
        openChangesOnlyFileButton.setMaxWidth(Double.MAX_VALUE);
        openChangesOnlyFileButton.setStyle("-fx-background-color: #cccccc; -fx-text-fill: black; -fx-font-size: 14px; -fx-background-radius: 8;");

        openChangesOnlyFileButton.setOnAction(e -> {
            try {
                Util.openFile(new File(changesOnlyFileName));
            } catch (Exception ex) {
                showAlert(ex.getMessage(), Alert.AlertType.ERROR);
                log(ex.getMessage());
            }
        });

        openFileButton.setMinWidth(180);
        openChangesOnlyFileButton.setMinWidth(180);

        HBox buttonsRow = new HBox(10, openFileButton, orLabel, openChangesOnlyFileButton);
        buttonsRow.setAlignment(Pos.CENTER);
        buttonsRow.setPadding(new Insets(5, 0, 5, 0));



        final VBox layout = new VBox(15, fileSection, submitButton, logPane, buttonsRow, creditLabel);
        layout.setPadding(new Insets(20));
        layout.setStyle("-fx-font-family: 'Segoe UI';");


        browseLabButton.setOnAction(e -> {
            labDataFile = promptSelectFile(stage);
            if (labDataFile != null) labDataField.setText(labDataFile.getAbsolutePath());
        });

        browseIndexButton.setOnAction(e -> {
            indexFile = promptSelectFile(stage);
            if (indexFile != null) indexFileField.setText(indexFile.getAbsolutePath());
        });

        submitButton.setOnAction(e -> {
            if (labDataFile != null && indexFile != null) {
                submitButton.setDisable(true);
                submitButton.setText("Processing...");
                openFileButton.setDisable(true);
                openFileButton.setStyle("-fx-background-color: #cccccc; -fx-text-fill: black; -fx-font-size: 14px; -fx-background-radius: 8;");

                final Task<Void> task = new Task<>() {
                    @Override
                    protected Void call() {
                        log("Reading lab data file...");
                        try {
                            readLabDataFile(labDataFile);
                            tryUpdateIndexFile(indexFile);
                            log("Completed processing. Updated file written to: " + editedIndexFileName);
                        } catch (Exception ex) {
                            log("Error: " + ex.getMessage());
                            showAlert(ex.getMessage(), Alert.AlertType.WARNING);
                        }
                        return null;
                    }

                    @Override
                    protected void succeeded() {
                        resetButton();
                        Platform.runLater(() -> {
                            openFileButton.setDisable(false);
                            openFileButton.setStyle("-fx-background-color: #038262; -fx-text-fill: white; -fx-font-size: 14px; -fx-background-radius: 8;");
                            openFileButton.setOnMouseEntered(e ->
                                    openFileButton.setStyle("-fx-font-size: 14px; -fx-background-color: #0a8c6c; -fx-text-fill: white; -fx-background-radius: 8;")
                            );
                            openFileButton.setOnMouseExited(e ->
                                    openFileButton.setStyle("-fx-font-size: 14px; -fx-background-color: #038262; -fx-text-fill: white; -fx-background-radius: 8;")
                            );
                        });
                    }

                    @Override
                    protected void failed() {
                        resetButton();
                    }

                    @Override
                    protected void cancelled() {
                        resetButton();
                    }
                };

                new Thread(task).start();
            } else {
                showAlert("Please select both files.", Alert.AlertType.WARNING);
            }
        });

        stage.setScene(new Scene(layout, 700, 500));
        stage.show();
    }


    // File Picker UI
    private File promptSelectFile(Stage stage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select Excel File");

        File selectedFile = fileChooser.showOpenDialog(stage);

        if (selectedFile != null) {
            log("File selected: " + selectedFile.getAbsolutePath());
        }

        return selectedFile;
    }

    private void tryUpdateIndexFile(File file) throws Exception {
        try (final FileInputStream fis = new FileInputStream(file);
             final Workbook workbook = new XSSFWorkbook(fis);
             final Workbook changesOnlyWorkBook = new XSSFWorkbook();
             ) {

            final Sheet changesOnlySheet = changesOnlyWorkBook.createSheet("Changes Only");
            final Calendar calendar = new GregorianCalendar();
            final Sheet indexSheet = workbook.getSheetAt(INDEX_FILE_DATA_SHEET);
            final Row sampleNameRow = indexSheet.getRow(INDEX_FILE_SAMPLES_ROW);
            final Row dateNameRow = indexSheet.getRow(INDEX_FILE_DATES_ROW);
            final List<String> seenTypes = new ArrayList<>();
            
            int updatedLocationsAmount = 0;

            // Loop though all cells in the index file
            for (Cell cell : sampleNameRow) {
                if (cell.getCellType() == CellType.STRING && !seenTypes.contains(cell.getStringCellValue())) {
                    seenTypes.add(cell.getStringCellValue());

                    final String currId = tryToMatchLocationID(cell.getStringCellValue());

                    if (labTestDataByLocationId.containsKey(currId)) {
                        final LocationTestData locationTestData = labTestDataByLocationId.get(currId);
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
                                issue = "Error reading date in column " + NumberTranslator.columnIndexToExcelLetter(currColIndex);
                            }
                        }

                        if (issue != null) {
                            log(currId + ": FAILED - " + issue);
                        } else if (needsUpdating) {
                            createNewCol(indexSheet, currColIndex);

                            final Cell newDateCell = dateNameRow.getCell(currColIndex);
                            newDateCell.setCellValue(newestDate);

                            for (Row row : indexSheet) {
                                final Cell paramCell = row.getCell(0);

                                if (paramCell != null && paramCell.getCellType() == CellType.STRING) {
                                    final String paramId = paramCell.getStringCellValue();
                                    final String data = locationTestData.getDataByParam(paramId);

                                    // Write data to index file
                                    writeToCell(row.getCell(currColIndex), data);

                                    // Write data to changes only sheet
                                    writeToCell(changesOnlySheet.createRow(row.getRowNum()).createCell(updatedLocationsAmount), data);
                                }
                            }

                            labTestDataByLocationId.remove(currId);

                            log(currId + ": COMPLETED - added new column at index " + NumberTranslator.columnIndexToExcelLetter(currColIndex) + ".");

                            updatedLocationsAmount++;

                        } else {
                            log(currId + ": SKIPPED - index file already up to date.");
                        }
                    }
                }
            }

            new File("/output").mkdirs();

            try (FileOutputStream fos = new FileOutputStream(editedIndexFileName)) {
                workbook.write(fos);
                log("Updated index file written to: " + editedIndexFileName);
            }
            
            try (FileOutputStream fos = new FileOutputStream(changesOnlyFileName)) {
                changesOnlyWorkBook.write(fos); // line 320
                log("Changes only file written to: " + changesOnlyFileName);
            }

        } catch (Exception ex) {
            ex.printStackTrace();
            log("Error: " + ex.getMessage());
            throw new Exception(ex.getMessage());
        }
    }

    // Location IDs in the index file are slightly different
    // Often they are longer in index files
    private String tryToMatchLocationID(String possibleId) {
        possibleId = possibleId.replace(" ", "").toLowerCase();

        for (String id : labTestDataByLocationId.keySet()) {
            if (possibleId.contains(id.replace(" ", "").toLowerCase())) {
                return id;
            }
        }

        return possibleId;
    }

    // Reads the lab data file and populates the labTestDataByLocationId map
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
                labTestDataByLocationId.put(sampleId.toUpperCase(), locationTestData);
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

                    // If the value is empty, we need to keep looking down the column
                    // This is for cases where data is not aligned due to different tests types
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
            throw new Exception("The lab data file is not in the expected format. ");
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

    // writes as a string or numeric value
    private void writeToCell(Cell cell, String value) {
        if (value == null || value.trim().isEmpty()) {
            cell.setBlank();
            return;
        }

        try {
            // fix some things copied as strings
            final double numericValue = Double.parseDouble(value);
            cell.setCellValue(numericValue);
        } catch (NumberFormatException e) {
            cell.setCellValue(value);
        }
    }

    // Creates a new column, coping the style from the left cells
    // Try to be consistent with the index file
    private void createNewCol(Sheet sheet, int colIndex) {
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

    private void log(String message) {
        if (Platform.isFxApplicationThread()) {
            logArea.appendText(message + "\n");
        } else {
            Platform.runLater(() -> logArea.appendText(message + "\n"));
        }
    }

    private void showAlert(String message, Alert.AlertType type) {
        Platform.runLater(() -> {
            Alert alert = new Alert(type, message, ButtonType.OK);
            alert.setTitle(APP_NAME);
            alert.showAndWait();
        });
    }

    private void resetButton() {
        Platform.runLater(() -> {
            submitButton.setDisable(false);
            submitButton.setText("Process and Update Index File");
        });
    }
}
