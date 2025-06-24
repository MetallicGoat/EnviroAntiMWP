package me.metallicgoat.datacombiner;

public class NumberTranslator {

    public static String columnIndexToExcelLetter(int column) {
        StringBuilder letter = new StringBuilder();
        while (column >= 0) {
            int remainder = column % 26;
            letter.insert(0, (char) (remainder + 'A'));
            column = (column / 26) - 1;
        }
        return letter.toString();
    }
}
