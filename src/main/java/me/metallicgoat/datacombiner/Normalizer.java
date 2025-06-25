package me.metallicgoat.datacombiner;

public class Normalizer {
    public static String normalize(String input) {
        if (input == null) return "";
        return input
                .toLowerCase()
                .replaceAll("\\b(cont'd|continued|duplicate)\\b", "")
                .replaceAll("[^a-z0-9]", "")
                .trim();
    }
}
