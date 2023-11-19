package org.julianjiang.javafx.utils;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;

import static org.julianjiang.javafx.Constants.SHEET_NAME_SPLIT;

public class CustomKeyComparator implements Comparator<String> {

    private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    @Override
    public int compare(String key1, String key2) {
        int dateCompare = compareDate(key1, key2);
        if (dateCompare != 0) {
            return dateCompare;
        }
        if (key1.length() == 10 || key2.length() == 10) {
            return dateCompare;
        }
        int categoryCompare = key1.substring(key1.lastIndexOf(SHEET_NAME_SPLIT) + 3).compareTo(key2.substring(key2.indexOf(SHEET_NAME_SPLIT) + 3));
        return categoryCompare;
    }

    private int compareDate(String key1, String key2) {
        try {
            LocalDate date1 = LocalDate.parse(key1.substring(0, 10), formatter);
            LocalDate date2 = LocalDate.parse(key2.substring(0, 10), formatter);
            return date1.compareTo(date2);
        } catch (Exception e) {
            return 0;
        }
    }
}