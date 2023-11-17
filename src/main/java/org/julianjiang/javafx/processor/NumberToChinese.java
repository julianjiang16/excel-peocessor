package org.julianjiang.javafx.processor;

import java.util.Formatter;

public class NumberToChinese {
    private static final String[] CHINESE_NUMBERS = {"零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"};
    private static final String[] CHINESE_UNITS = {"", "拾", "佰", "仟"};
    private static final String[] CHINESE_GROUP_UNITS = {"", "万", "亿"};
    private static final String CHINESE_GROUP_SEPARATOR = ".";

    public static String convert(double number) {
        StringBuilder sb = new StringBuilder();
        Formatter formatter = new Formatter(sb);
        String numberStr = Double.toString(number);
        int decimalIndex = numberStr.indexOf(CHINESE_GROUP_SEPARATOR);
        String decimalPart = decimalIndex != -1 ? numberStr.substring(decimalIndex) : "";
        String integerPart = decimalIndex != -1 ? numberStr.substring(0, decimalIndex) : numberStr;
        int groupUnitIndex = integerPart.indexOf(CHINESE_GROUP_SEPARATOR);
        String groupUnit = groupUnitIndex != -1 ? integerPart.substring(groupUnitIndex) : "";
        String integerPartStr = groupUnitIndex != -1 ? integerPart.substring(0, groupUnitIndex) : integerPart;
        int lastThreeDigitsIndex = integerPartStr.length() >= 3 ? integerPartStr.length() - 3 : 0;
        String lastThreeDigitsStr = integerPartStr.substring(lastThreeDigitsIndex);
        String remainingDigitsStr = integerPartStr.substring(0, lastThreeDigitsIndex);
        int lastDigitIndex = lastThreeDigitsStr.length() >= 1 ? lastThreeDigitsStr.length() - 1 : 0;
        String lastDigitStr = lastThreeDigitsStr.substring(lastDigitIndex);
        String remainingDigitsStr2 = lastThreeDigitsStr.substring(0, lastDigitIndex);
        sb.append(CHINESE_NUMBERS[Integer.parseInt(lastDigitStr)]);
        sb.append(CHINESE_UNITS[Integer.parseInt(remainingDigitsStr)]);
        sb.append(CHINESE_GROUP_UNITS[Integer.parseInt(remainingDigitsStr2)]);
        sb.append("万");
        sb.append(CHINESE_GROUP_UNITS[Integer.parseInt(groupUnit)]);
        if (!decimalPart.equals("")) {
            sb.append(".");
            for (int i = 0; i < decimalPart.length(); i++) {
                int digit = Integer.parseInt(String.valueOf(decimalPart.charAt(i)));
                sb.append(CHINESE_NUMBERS[digit]);
                if (i < decimalPart.length() - 1) {
                    sb.append(CHINESE_UNITS[9 - i]);
                } else {
                    sb.append("点");
                }
            }
        } else {
            sb.append("整");
        }
        return formatter.toString();
    }

    public static void main(String[] args) {
        System.err.println(convert(12064.89));
    }
}
