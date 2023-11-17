package org.julianjiang.javafx.utils;

import com.google.common.collect.Lists;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Map;

public class FormulaUtils {
    private static ScriptEngineManager manager = new ScriptEngineManager();
    private static ScriptEngine engine = manager.getEngineByName("JavaScript");
    private static DecimalFormat df = new DecimalFormat("#.00");

    static {
        df.setRoundingMode(RoundingMode.HALF_UP);
    }

    private static final String ADD = "\\+";
    private static final String REDUCE = "-";
    private static final String RIDE = "\\*";
    private static final String DIVIDE = "/";

    public static Object analysisFormula(String formula) throws ScriptException {
        return Double.valueOf(df.format(engine.eval(formula)));
    }


    public static String convert2Formula(String sourceStr, Map<String, Object> replaceMap) {

        String join = "";
        final ArrayList<String> field = Lists.newArrayList();
        if (sourceStr.indexOf(ADD) > -1) {
            field.addAll(Arrays.asList(sourceStr.split(ADD)));
            join = ADD;
        } else if (sourceStr.indexOf(REDUCE) > -1) {
            field.addAll(Arrays.asList(sourceStr.split(REDUCE)));
            join = REDUCE;
        } else if (sourceStr.indexOf(RIDE) > -1) {
            field.addAll(Arrays.asList(sourceStr.split(RIDE)));
            join = RIDE;
        } else if (sourceStr.indexOf(DIVIDE) > -1) {
            field.addAll(Arrays.asList(sourceStr.split(DIVIDE)));
            join = DIVIDE;
        }

        for (int i = 0; i < field.size(); i++) {
            field.set(i, String.valueOf(replaceMap.getOrDefault(field.get(i).trim(), field.get(i))));
        }
        final String formula = String.join(join, field);
        return formula;
    }
}
