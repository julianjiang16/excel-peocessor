package org.julianjiang.javafx;

import javafx.scene.control.TextField;
import lombok.Data;
import org.julianjiang.javafx.model.ExcelTemplate;

import java.io.File;
import java.util.List;
import java.util.Map;

@Data
public class Context {

    public Context(ExcelTemplate excelTemplate) {
        this.excelTemplate = excelTemplate;
    }

    // excel数据
    List<Map<String, Object>> data;

    // 分单字段
    List<String> allocation;

    boolean typeFlag;

    @Deprecated
    File templateFile;

    // 前N行
    @Deprecated
    int preRows;

    // 后N行
    @Deprecated
    int lastRows;

    String outputPath;

    @Deprecated
    TextField preText;
    @Deprecated
    TextField lastText;

    String lastFilePath;

    ExcelTemplate excelTemplate;
}
