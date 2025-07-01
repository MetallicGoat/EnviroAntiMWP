package me.metallicgoat.datacombiner;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import me.metallicgoat.datacombiner.handlers.IndexFileHandler;
import me.metallicgoat.datacombiner.handlers.LabDataFileHandler;
import me.metallicgoat.datacombiner.util.Helpers;
import me.metallicgoat.datacombiner.util.Util;

import java.io.File;

public class UserInterface {

    private File labDataFile;
    private File indexFile;
    private TextArea logArea;
    private Button submitButton;
    private Button openFileButton;
    private Button openChangesOnlyFileButton;

    public void start(Stage stage) {
        stage.setTitle("Index Updater");

        // UI components for selecting the lab and index files
        final Label labDataLabel = new Label("Lab Data File:");
        final TextField labDataField = new TextField();
        labDataField.setPromptText("No file selected");
        labDataField.setEditable(false);

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

        Label creditLabel = new Label("Created by Christian Azzam and Braydon Affleck · Making life less painful");
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
                Helpers.openFile(new File(Util.EDITED_INDEX_FILE_PATH));
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
                Helpers.openFile(new File(Util.CHANGES_ONLY_FILE_PATH));
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
            labDataFile = Util.getUI().promptSelectFile(stage, "Lab Data");
            if (labDataFile != null) labDataField.setText(labDataFile.getAbsolutePath());
        });

        browseIndexButton.setOnAction(e -> {
            indexFile = promptSelectFile(stage, "Index");
            if (indexFile != null) indexFileField.setText(indexFile.getAbsolutePath());
        });

        submitButton.setOnAction(e -> {
            if (labDataFile != null && indexFile != null) {
                submitButton.setDisable(true);
                submitButton.setText("Processing...");
                setFileOpenButtonEnabled(openFileButton, false);
                setFileOpenButtonEnabled(openChangesOnlyFileButton, false);

                final Task<Void> task = new Task<>() {
                    @Override
                    protected Void call() {
                        log("Reading lab data file...");
                        try {
                            LabDataFileHandler.readAndStoreData(labDataFile);
                            IndexFileHandler.tryUpdateIndexFile(indexFile);
                            log("Completed processing.");
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
                            setFileOpenButtonEnabled(openFileButton, true);
                            setFileOpenButtonEnabled(openChangesOnlyFileButton, true);
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


    private void setFileOpenButtonEnabled(Button button, boolean enabled) {
        if (enabled) {
            button.setDisable(false);
            button.setStyle("-fx-background-color: #038262; -fx-text-fill: white; -fx-font-size: 14px; -fx-background-radius: 8;");
            button.setOnMouseEntered(e ->
                    button.setStyle("-fx-font-size: 14px; -fx-background-color: #0a8c6c; -fx-text-fill: white; -fx-background-radius: 8;")
            );
            button.setOnMouseExited(e ->
                    button.setStyle("-fx-font-size: 14px; -fx-background-color: #038262; -fx-text-fill: white; -fx-background-radius: 8;")
            );
        } else {
            button.setDisable(true);
            button.setStyle("-fx-background-color: #cccccc; -fx-text-fill: black; -fx-font-size: 14px; -fx-background-radius: 8;");
        }
    }

    // File Picker UI
    private File promptSelectFile(Stage stage, String fileName) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Select " + fileName + " File");

        File selectedFile = fileChooser.showOpenDialog(stage);

        if (selectedFile != null) {
            log("File selected: " + selectedFile.getAbsolutePath());
        }

        return selectedFile;
    }


    public void log(String message) {
        if (Platform.isFxApplicationThread()) {
            logArea.appendText(message + "\n");
        } else {
            Platform.runLater(() -> logArea.appendText(message + "\n"));
        }
    }

    public void showAlert(String message, Alert.AlertType type) {
        Platform.runLater(() -> {
            Alert alert = new Alert(type, message, ButtonType.OK);
            alert.setTitle(Util.APP_NAME);
            alert.showAndWait();
        });
    }

    public void resetButton() {
        Platform.runLater(() -> {
            submitButton.setDisable(false);
            submitButton.setText("Process and Update Index File");
        });
    }
}
