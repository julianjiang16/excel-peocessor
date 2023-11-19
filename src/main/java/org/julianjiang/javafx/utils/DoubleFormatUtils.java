package org.julianjiang.javafx.utils;

import java.text.DecimalFormat;

public class DoubleFormatUtils {

    public static String format(double value) {
        DecimalFormat df = new DecimalFormat("#.##");
        String formatStr = df.format(value);
        if (formatStr.endsWith(".00")) {
            formatStr = formatStr.substring(0, formatStr.length() - 3);
        }
        return formatStr;
    }
}
