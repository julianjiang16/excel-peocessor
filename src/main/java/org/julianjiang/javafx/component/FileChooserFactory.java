package org.julianjiang.javafx.component;

import javafx.geometry.Insets;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.math3.util.Pair;
import org.julianjiang.javafx.model.Context;
import org.julianjiang.javafx.utils.ExcelUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class FileChooserFactory {

    Font font;
    String fontTemplate;
    Label label;

    public FileChooserFactory(Font font, Label label, String fontTemplate) {
        this.font = font;
        this.fontTemplate = String.format(fontTemplate, font.getName());
        this.label = label;
        this.label.setStyle(fontTemplate);
        this.label.setPrefWidth(300);
    }

    public void setLabel(Label label) {
        this.label = label;
    }

    public HBox buildFileChooser(Stage primaryStage, ListView comboBox, Context context, boolean pathFlag) {
        Button button = new Button(pathFlag ? "选择文件夹" : "选择文件");
        button.setStyle(fontTemplate);

        TextField textField = new TextField();
        textField.setPrefWidth(300);
        textField.setEditable(false);
        // 设置文本框样式为灰色
        textField.setStyle("-fx-background-color: lightgray");

        button.setOnAction(e -> {
            if (pathFlag) {
                buildPahChooser(primaryStage, context, textField);
                return;
            }
            buildFileChooser(primaryStage, comboBox, context, textField);
        });
        final HBox hBoxFile = new HBox(10);
        hBoxFile.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        hBoxFile.getChildren().addAll(label, textField, button);
        return hBoxFile;
    }

    private void buildFileChooser(Stage primaryStage, ListView comboBox, Context context, TextField textField) {

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Choose File");
        if (StringUtils.isNotBlank(context.getLastFilePath())) {
            fileChooser.setInitialDirectory(new File(context.getLastFilePath()));
        } else {
            fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
        }
        FileChooser.ExtensionFilter excelFilter = new FileChooser.ExtensionFilter("excel文件 (*.xls, *.xlsx)", "*.xls", "*.xlsx");
        fileChooser.getExtensionFilters().add(excelFilter);
        File selectedFile = fileChooser.showOpenDialog(primaryStage);
        if (selectedFile != null) {
            context.setLastFilePath(selectedFile.getParent());
            Alert alert = AlertComponent.buildAlert("请稍等...", "程序正在处理，请稍等...");
            InputStream inputStream = null;
            try {
                alert.show();
                inputStream = new FileInputStream(selectedFile);
                if (!Objects.isNull(comboBox)) {
                    Pair<ArrayList<String>, List<Map<String, Object>>> dataPair = ExcelUtils.readExcel(inputStream);
                    List<String> strings = dataPair.getFirst();
                    comboBox.getItems().clear();
                    comboBox.getItems().addAll(strings);
                    comboBox.getSelectionModel().select(0);
                    context.setData(dataPair.getSecond());
                } else {
                    final List<String> replaceNames = ExcelUtils.getReplaceNames(inputStream);
                    // 读取模板文件，找出需要替换的预设字段规则
                    context.setReplaceNames(replaceNames);
                    context.getExcelTemplate().setTemplateFile(selectedFile);
                }
                textField.setText(selectedFile.getName());
            } catch (Exception ex) {
                ex.printStackTrace();
            } finally {
                if (inputStream != null) {
                    try {
                        inputStream.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if (alert.isShowing()) {
                    alert.close();
                }
            }
        }
    }

    private void buildPahChooser(Stage primaryStage, Context context, TextField textField) {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("选择文件夹");

        if (StringUtils.isNotBlank(context.getLastFilePath())) {
            directoryChooser.setInitialDirectory(new File(context.getLastFilePath()));
        } else {
            directoryChooser.setInitialDirectory(new File(System.getProperty("user.home")));
        }
        File selectedDirectory = directoryChooser.showDialog(primaryStage);
        if (selectedDirectory != null) {
            context.setLastFilePath(selectedDirectory.getParent());
            context.setOutputPath(selectedDirectory.getAbsolutePath());
            textField.setText(selectedDirectory.getAbsolutePath());
        }
    }
}
