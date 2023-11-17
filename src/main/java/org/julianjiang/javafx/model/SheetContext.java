package org.julianjiang.javafx.model;

import com.google.common.collect.Maps;
import lombok.Data;

import java.util.Map;

@Data
public class SheetContext {

    // 日期
    String saleDate;
    // 点位
    String pointPlace;

    String typeName;

    int typeCnt;

    float typeAmount;

    int cnt;

    float amount;

    // key 带后缀
    Map<String,Object> replaceMap = Maps.newHashMap();

}
