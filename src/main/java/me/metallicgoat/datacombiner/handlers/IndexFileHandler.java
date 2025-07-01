package me.metallicgoat.datacombiner.handlers;

import me.metallicgoat.datacombiner.util.Helpers;
import me.metallicgoat.datacombiner.util.LocationTestData;
import me.metallicgoat.datacombiner.util.Util;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class IndexFileHandler {


    public static void tryUpdateIndexFile(File file) throws Exception {
        try (final FileInputStream fis = new FileInputStream(file);
             final Workbook workbook = new XSSFWorkbook(fis);
             final Workbook changesOnlyWorkBook = new XSSFWorkbook()
        ) {
            final Sheet changesOnlySheet = changesOnlyWorkBook.createSheet("Changes Only");
            final Calendar currentCalender = new GregorianCalendar();
            final Calendar newestCalender = new GregorianCalendar();

            Sheet indexSheet = null;
            boolean foundTabGrf = false;

            // TODO no way in hell this handles all cases
            // Look for the sheet immediately after "Tab-GRF" or "TC-GRF"
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                final String sheetName = workbook.getSheetName(i).toLowerCase();
                if (foundTabGrf) {
                    indexSheet = workbook.getSheetAt(i);
                    break;
                }

                if (sheetName.equals("tab-grf") || sheetName.equals("tc-grf")) {
                    foundTabGrf = true;
                }
            }


            if (indexSheet == null) {
                throw new Exception("Could not find sheet after 'Tab-GRF'.");
            }


            // Find the significant rows in this specific file (DATE ROW STUFF)
            locateSignificantRows(indexSheet);


            Util.getUI().log("Updating index and change files...");

            final Row sampleNameRow = indexSheet.getRow(Util.INDEX_FILE_SAMPLES_ROW);
            final Row dateNameRow = indexSheet.getRow(Util.INDEX_FILE_DATES_ROW);

            // 1. Snapshot of all lab IDs normalized BEFORE processing
            Set<String> allLabIdsNormalized = new HashSet<>();
            for (String id : Util.getStoredLabDataLocations()) {
                allLabIdsNormalized.add(Helpers.normalizeSampleId(id));
            }

            // To keep track of which lab IDs got matched this run
            Set<String> matchedLabIdsNormalized = new HashSet<>();

            int updatedLocationsAmount = 0;

            // TODO: I might have fixed the concurrent modification
            List<Cell> cells = new ArrayList<>();
            for (Cell cell : sampleNameRow) {
                cells.add(cell);
            }

            Set<String> alreadyHandledIds = new HashSet<>();

            // 2. Process each index cell to add new data columns if needed
            for (Cell cell : cells) {
                if (cell.getCellType() == CellType.STRING) {
                    String currRawId = cell.getStringCellValue();
                    String normalizedCurrId = Helpers.normalizeSampleId(currRawId);

                    if (alreadyHandledIds.contains(normalizedCurrId)) {
                        continue; // skip duplicates and cont'd versions
                    }

                    alreadyHandledIds.add(normalizedCurrId);

                    // Find matching lab ID key for this index ID
                    String matchedLabId = null;
                    for (String labId : Util.getStoredLabDataLocations()) {
                        if (Helpers.normalizeSampleId(labId).equals(normalizedCurrId)) {
                            matchedLabId = labId;
                            break;
                        }
                    }

                    if (matchedLabId != null) {
                        final LocationTestData locationTestData = Util.getStoredLabDataByLocationId(matchedLabId);
                        final Date newestDate = locationTestData.dataDate;
                        newestCalender.setTime(newestDate);

                        int currColIndex = cell.getColumnIndex();
                        int curYear = 0;
                        boolean keepLooking = true;
                        boolean needsUpdating = true;
                        String issue = null;

                        while (keepLooking) {
                            final Cell dateCell = dateNameRow.getCell(currColIndex);
                            if (dateCell == null) break;

                            try {
                                final Date date = dateCell.getDateCellValue();
                                currentCalender.setTime(date);
                                final int year = currentCalender.get(Calendar.YEAR);

                                if (curYear <= year) {
                                    if (currentCalender.get(Calendar.YEAR) == newestCalender.get(Calendar.YEAR) &&
                                            currentCalender.get(Calendar.MONTH) == newestCalender.get(Calendar.MONTH)) {
                                        needsUpdating = false;
                                        break;
                                    }
                                    curYear = year;
                                    currColIndex++;
                                } else {
                                    keepLooking = false;
                                }
                            } catch (Exception ex) {

                                // TODO Broken warning after refactor
                                // 'curRow' is encapsulated now.

//                                String cellInfo;
//
//                                try {
//                                    Cell debugCell = indexSheet.getRow(currRow).getCell(currColIndex);
//                                    if (debugCell == null) {
//                                        cellInfo = "Cell is null.";
//                                    } else {
//                                        cellInfo = "Cell content: '" + debugCell.toString() + "', type: " + debugCell.getCellType();
//                                    }
//                                } catch (Exception nestedEx) {
//                                    cellInfo = "Unable to inspect cell.";
//                                }
//
//                                keepLooking = false;
//                                issue = "Error reading date in column " + Helpers.columnIndexToExcelLetter(currColIndex)
//                                        + " at row " + (currRow + 1) + ". " + cellInfo;
//
//                                ex.printStackTrace(); // Optional: disable later
                            }
                        }

                        if (issue != null) {
                            Util.getUI().log(currRawId + ": FAILED - " + issue);

                        } else if (needsUpdating) {
                            Helpers.createNewColWithCopiedStyles(indexSheet, currColIndex);

                            final Cell newDateCell = dateNameRow.getCell(currColIndex);
                            newDateCell.setCellValue(newestDate);

                            // Optional: apply date style if you have one
                            // newDateCell.setCellStyle(dateCellStyle);

                            for (Row row : indexSheet) {
                                final Cell paramCell = row.getCell(0);
                                if (paramCell != null && paramCell.getCellType() == CellType.STRING) {
                                    String paramId = paramCell.getStringCellValue();
                                    String data = locationTestData.getDataByParam(paramId);

                                    Helpers.writeToCell(row.getCell(currColIndex), data);

                                    Row changesOnlyRow = changesOnlySheet.getRow(row.getRowNum());
                                    if (changesOnlyRow == null) {
                                        changesOnlyRow = changesOnlySheet.createRow(row.getRowNum());
                                    }
                                    Helpers.writeToCell(changesOnlyRow.createCell(updatedLocationsAmount), data);
                                }
                            }

                            // Add Title (Changes only file)
                            Row changesOnlyTitleRow = changesOnlySheet.getRow(Util.INDEX_FILE_SAMPLES_ROW);
                            if (changesOnlyTitleRow == null) {
                                changesOnlyTitleRow = changesOnlySheet.createRow(Util.INDEX_FILE_SAMPLES_ROW);
                            }
                            changesOnlyTitleRow.createCell(updatedLocationsAmount).setCellValue(currRawId);

                            // Add date (Changes only file)
                            Row changesOnlyDateRow = changesOnlySheet.getRow(Util.INDEX_FILE_DATES_ROW);
                            if (changesOnlyDateRow == null) {
                                changesOnlyDateRow = changesOnlySheet.createRow(Util.INDEX_FILE_DATES_ROW);
                            }
                            Cell dateCell = changesOnlyDateRow.createCell(updatedLocationsAmount);
                            dateCell.setCellValue(newestDate);
                            // Optionally apply dateCellStyle here too

                            matchedLabIdsNormalized.add(normalizedCurrId);

                            Util.getUI().log(currRawId + ": COMPLETED - added new column at index " + Helpers.columnIndexToExcelLetter(currColIndex) + ".");

                            updatedLocationsAmount++;
                        } else {
                            Util.getUI().log(currRawId + ": SKIPPED - index file already up to date.");
                            matchedLabIdsNormalized.add(normalizedCurrId);
                        }
                    }
                }
            }

            // 3. Now find index IDs that have no matching lab data
            Set<String> normalizedSampleIdsInIndex = new HashSet<>();
            List<String> unmatchedIndexIds = new ArrayList<>();

            for (Cell cell : sampleNameRow) {
                if (cell.getCellType() == CellType.STRING) {
                    final String rawId = cell.getStringCellValue();

                    if (Util.isNonWellLabel(rawId)) {
                        continue;
                    }

                    final String normalizedId = Helpers.normalizeSampleId(rawId);

                    if (!normalizedSampleIdsInIndex.contains(normalizedId)) {
                        normalizedSampleIdsInIndex.add(normalizedId);

                        if (!allLabIdsNormalized.contains(normalizedId)) {
                            unmatchedIndexIds.add(rawId);
                        }
                    }
                }
            }

            if (!unmatchedIndexIds.isEmpty()) {
                Util.getUI().log("Indices which were not updated:");

                for (String id : unmatchedIndexIds) {
                    Util.getUI().log(id);
                }

            } else {
                Util.getUI().log("All wells in index file updated.");
            }

            // 4. Find lab IDs that were never matched to index (missed wells)
            List<String> unmatchedLabIds = new ArrayList<>();
            for (String labId : Util.getStoredLabDataLocations()) {
                String normalizedLabId = Helpers.normalizeSampleId(labId);
                if (!matchedLabIdsNormalized.contains(normalizedLabId)) {
                    unmatchedLabIds.add(labId);
                }
            }

            // 5. Create output directory if it doesn't exist
            new File("output").mkdirs();

            // 6. Write output files
            try (FileOutputStream fos = new FileOutputStream(Util.EDITED_INDEX_FILE_PATH)) {
                workbook.write(fos);
                Util.getUI().log("Updated index file written to: " + Util.EDITED_INDEX_FILE_PATH);
            }

            try (FileOutputStream fos = new FileOutputStream(Util.CHANGES_ONLY_FILE_PATH)) {
                changesOnlyWorkBook.write(fos);
                Util.getUI().log("Changes only file written to: " + Util.CHANGES_ONLY_FILE_PATH);
            }

        } catch (Exception ex) {
            ex.printStackTrace();
            Util.getUI().log("Error: " + ex.getMessage());
            throw new Exception(ex.getMessage());
        }
    }

    private static void locateSignificantRows(Sheet indexSheet ) {
        final FormulaEvaluator evaluator = indexSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int dateColumnGuess = 1; // Start with column B (0-based index)
        boolean dateFound = false;

        while (!dateFound && dateColumnGuess < 20) {
            // Skip hidden columns
            if (indexSheet.isColumnHidden(dateColumnGuess)) {
                // log("Skipping hidden column: " + dateColumnGuess);
                dateColumnGuess++;
                continue;
            }

            for (int rowIndex = 2; rowIndex < 10; rowIndex++) {
                Row row = indexSheet.getRow(rowIndex);
                if (row == null) continue;

                Cell cell = row.getCell(dateColumnGuess);
                String address = new CellAddress(rowIndex, dateColumnGuess).formatAsString();

                // Diagnostic code to see why date finder is failing
                // if (cell != null) {
                //     log("Checking cell " + address + ": " + cell.toString() + " type: " + cell.getCellType());
                // } else {
                //     log("Checking cell " + address + ": null");
                // }

                if (cell != null) {
                    CellType cellType = cell.getCellType();

                    if (cellType == CellType.NUMERIC) {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            Util.INDEX_FILE_DATES_ROW = rowIndex;
                            Util.INDEX_FILE_SAMPLES_ROW = rowIndex - 1;
                            dateFound = true;
                            break;
                        }
                    } else if (cellType == CellType.FORMULA) {
                        CellValue evaluatedValue = evaluator.evaluate(cell);
                        if (evaluatedValue != null
                                && evaluatedValue.getCellType() == CellType.NUMERIC
                                && DateUtil.isValidExcelDate(evaluatedValue.getNumberValue())) {

                            // This is a date formula cell
                            Util.INDEX_FILE_DATES_ROW = rowIndex;
                            Util.INDEX_FILE_SAMPLES_ROW = rowIndex - 1;
                            dateFound = true;
                            break;
                        }
                    }
                }
            }

            if (!dateFound) {
                dateColumnGuess++;
            }
        }

        if (dateFound) {
            Util.getUI().log("Found Date and Sample rows.");
        }

        if (!dateFound) {
            Util.getUI().log("Failed to guess date and samples row from index file.");
            Util.getUI().log("Falling back to hard coded values");
        }
    }
}
