package org.julianjiang.javafx.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PatternUtils {

    public static String getMatchedString(String input) {
        // 编译正则表达式
        Pattern pattern = Pattern.compile("#(.*?)#");
        // 创建 matcher 对象
        Matcher matcher = pattern.matcher(input);
        // 查找与给定模式匹配的输入序列的第一个子序列
        if (matcher.find()) {
            return matcher.group(1);
        } else {
            return null;
        }
    }
}
