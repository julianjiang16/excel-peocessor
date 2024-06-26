package org.julianjiang.javafx.model;

import javafx.scene.control.TextField;
import lombok.Data;

import java.io.File;

@Data
public class ExcelTemplate {

    // 前N行和最后N行 原样输出
    TextField preText;
    TextField lastText;

    File templateFile;

    // 列名所处的行号
    TextField titleText;
    // 如果选择了需要分类，则需要填写分类行样式的行号
    TextField typeText;

    // 明细数据样式行号
    TextField detailText;

    // 冗余
    int preRows;
    int lastRows;

    // 列名行号
    int titleRow;
    // 分类汇总行（样式）行号
    int typeRow;
    // 明细数据行（样式）行号
    int detailRow;



}
