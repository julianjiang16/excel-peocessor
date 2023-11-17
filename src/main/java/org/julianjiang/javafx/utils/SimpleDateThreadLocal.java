package org.julianjiang.javafx.utils;

import java.text.SimpleDateFormat;

public class SimpleDateThreadLocal {

    private static ThreadLocal<SimpleDateFormat> threadLocal = new ThreadLocal<>();

    public static SimpleDateFormat getSimpleDateFormat(String pattern) {
        SimpleDateFormat sdf = threadLocal.get();
        if (sdf == null) {
            sdf = new SimpleDateFormat(pattern);
            threadLocal.set(sdf);
        } else {
            sdf.applyPattern(pattern);
        }
        return sdf;
    }
}
