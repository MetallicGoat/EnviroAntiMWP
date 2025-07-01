package me.metallicgoat.datacombiner.util;

import me.metallicgoat.datacombiner.UserInterface;

import java.util.HashMap;
import java.util.Set;

public class Util {

    public static final String APP_NAME = "EnviroAntiMWP";

    public static final String EDITED_INDEX_FILE_PATH = "output/UpdatedIndex.xlsx";
    public static final String CHANGES_ONLY_FILE_PATH = "output/Changes.xlsx";

    public static int INDEX_FILE_SAMPLES_ROW = 4; // This is a fallback if search fails
    public static int INDEX_FILE_DATES_ROW = 5; // This is a fallback if search fails
    public static final int INDEX_FILE_DATA_SHEET = 1; // Fallback if no sheet has "tab-"


    private static UserInterface userInterface = new UserInterface();
    private static final HashMap<String, LocationTestData> labTestDataByLocationId = new HashMap<>();
    private static final Set<String> nonWellLabels = Set.of("PARAMETER", "Units", "UNITS", "ODWQS", "PWQO");

    public static void initUserInterface(UserInterface instance) {
        userInterface = instance;
    }

    public static UserInterface getUI() {
        return userInterface;
    }

    public static void storeLabData(String labId, LocationTestData data) {
        labTestDataByLocationId.put(labId, data);
    }

    public static Set<String> getStoredLabDataLocations() {
        return labTestDataByLocationId.keySet();
    }

    public static LocationTestData getStoredLabDataByLocationId(String labId) {
        return labTestDataByLocationId.get(labId);
    }

    public static boolean isNonWellLabel(String id) {
        return nonWellLabels.contains(id);
    }
}
