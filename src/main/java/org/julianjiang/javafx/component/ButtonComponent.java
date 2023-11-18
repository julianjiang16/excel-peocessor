package org.julianjiang.javafx.component;

import com.google.common.collect.Lists;
import javafx.application.Platform;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import javafx.stage.Stage;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.julianjiang.javafx.model.Context;
import org.julianjiang.javafx.processor.ExcelProcessor;

import javax.script.ScriptException;
import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import static org.julianjiang.javafx.utils.ExcelUtils.cacheCellStyle;

public class ButtonComponent {

    final ExecutorService executorService = Executors.newFixedThreadPool(1);

    public HBox buildButtonBox(ListView listView, Context context, Stage primaryStage) {
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

            if (!CollectionUtils.isEmpty(context.getAllocation())) {
                context.getAllocation().clear();
            } else {
                context.setAllocation(Lists.newArrayList());
            }
            context.getAllocation().addAll(selectedItems);

            // 设置模板前后保留行
            context.getExcelTemplate().setPreRows(getTextVal(context.getExcelTemplate().getPreText()));
            context.getExcelTemplate().setLastRows(getTextVal(context.getExcelTemplate().getLastText()));
            context.getExcelTemplate().setDetailRow(getTextVal(context.getExcelTemplate().getDetailText()));
            context.getExcelTemplate().setTitleRow(getTextVal(context.getExcelTemplate().getTitleText()));
            context.getExcelTemplate().setTypeRow(getTextVal(context.getExcelTemplate().getTypeText()));

            // 首先校验
            Pair<String, Boolean> validatePair = ExcelProcessor.validate(context);
            if (validatePair.getSecond()) {
                // 通过则弹出模板输出 路径选择框
                Alert alert = AlertComponent.buildAlert("处理中...", "正在拼命生成文件中...");
                try {
                    alert.show();
                    executorService.execute(() -> {
                        try {
                            ExcelProcessor.outputExcel(context);
                            Platform.runLater(() -> {
                                if (alert.isShowing()) {
                                    alert.close();
                                }
                            });
                            Platform.runLater(() -> AlertComponent.buildAlert("成功", "文件生成成功！！！").show());
                        } catch (IOException | ScriptException | InvalidFormatException e) {
                            e.printStackTrace();
                            Platform.runLater(() -> AlertComponent.buildAlert("错误", "文件生成失败！！！").show());
                        } finally {
                            cacheCellStyle.clear();
                        }
                    });
                } finally {
                    System.err.println("文件生成任务已提交！！！");
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
        cancel.setOnAction(event -> primaryStage.close());

        primaryStage.setOnCloseRequest(event -> {
            Platform.runLater(() -> executorService.shutdown());
        });

        hBox.getChildren().addAll(process, cancel);
        return hBox;
    }

    private int getTextVal(TextField field) {
        try {
            final Integer val = Integer.valueOf(field.getText());
            return val;
        } catch (Exception e) {
            return 0;
        }
    }
}
