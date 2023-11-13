package org.julianjiang.javafx.component;

import com.google.common.collect.Lists;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.math3.util.Pair;
import org.julianjiang.javafx.Context;
import org.julianjiang.javafx.processor.ExcelProcessor;

public class ButtonComponent {

    public HBox buildButtonBox(ListView listView, Context context) {
        HBox hBox = new HBox();
        hBox.setAlignment(Pos.CENTER);
        hBox.setPadding(new Insets(30, 0, 30, 20)); // 设置左边距
        hBox.setSpacing(100); // 设置组件间距
        Font f = Font.font("Arial", 20);
        Button process = new Button();
        process.setText("执行");
        process.setFont(f);

        process.setOnAction(actionEvent -> {
            // 设置分单条件
            MultipleSelectionModel<String> selectionModel = listView.getSelectionModel();
            ObservableList<String> selectedItems = selectionModel.getSelectedItems();
            selectedItems.stream().forEach(i -> {
                if (CollectionUtils.isEmpty(context.getAllocation())) {
                    context.setAllocation(Lists.newArrayList(i));
                } else {
                    context.getAllocation().add(i);
                }
            });

            // 设置模板前后保留行
            context.getExcelTemplate().setPreRows(Integer.valueOf(context.getExcelTemplate().getPreText().getText()));
            context.getExcelTemplate().setLastRows(Integer.valueOf(context.getExcelTemplate().getLastText().getText()));

            // 首先校验
            Pair<String, Boolean> validatePair = ExcelProcessor.validate(context);
            if (validatePair.getSecond()) {
                // 通过则弹出模板输出 路径选择框
                Alert alert = AlertComponent.buildAlert("处理中...", "正在拼命生成文件中...");
                try {
                    alert.show();
                    ExcelProcessor.outputExcel(context);
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    if (alert.isShowing()) {
                        alert.close();
                    }
                }
                return;
            }
            // 不通过则dialog
            Alert alert = AlertComponent.buildAlert("校验错误", validatePair.getFirst());
            alert.show();
        });
//
//        Button print = new Button();
//        print.setText("打印");
//        print.setFont(f);
//
//        print.setOnAction(event -> PrintProcessor.print(new File("C:\\Users\\Administrator\\Desktop\\打单模板 2.xlsx")));

        Button cancel = new Button();
        cancel.setText("关闭");
        cancel.setFont(f);
        hBox.getChildren().addAll(process, cancel);
        return hBox;
    }
}
