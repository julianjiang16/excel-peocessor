package org.julianjiang.javafx;

import javafx.scene.control.TextField;
import lombok.Data;

import java.io.File;
import java.util.List;
import java.util.Map;

@Data
public class Context {

    // excel数据
    List<Map<String, Object>> data;

    // 分单字段
    List<String> allocation;

    boolean typeFlag;

    File templateFile;

    // 前N行
    int preRows;

    // 后N行
    int lastRows;


    String outputPath;


    TextField preText;

    TextField lastText;

}
