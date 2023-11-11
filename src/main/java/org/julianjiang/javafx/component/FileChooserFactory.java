package org.julianjiang.javafx.component;

import javafx.geometry.Insets;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.julianjiang.javafx.processor.FileProcessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Objects;

public class FileChooserFactory {

    Font font;
    String fontTemplate;
    Label label;

    public FileChooserFactory(Font font, Label label) {
        this.font = font;
        String fontTemplate = "-fx-font: 18px \"%s\"; -fx-font-weight: bold;";
        this.fontTemplate = String.format(fontTemplate, font.getName());
        this.label = label;
        this.label.setStyle(fontTemplate);
        this.label.setPrefWidth(300);
    }

    public void setLabel(Label label) {
        this.label = label;
    }

    public HBox buildFileChooser(Stage primaryStage, ListView comboBox) {
        Button button = new Button("选择文件");
        button.setStyle(fontTemplate);

        TextField textField = new TextField();
        textField.setPrefWidth(300);
        textField.setEditable(false);
        // 设置文本框样式为灰色
        textField.setStyle("-fx-background-color: lightgray");

        if (!Objects.isNull(comboBox)) {
            button.setOnAction(e -> {
                FileChooser fileChooser = new FileChooser();
                fileChooser.setTitle("Choose File");
                fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
                FileChooser.ExtensionFilter excelFilter = new FileChooser.ExtensionFilter("excel文件 (*.xls, *.xlsx)", "*.xls", "*.xlsx");

                fileChooser.getExtensionFilters().add(excelFilter);
                File selectedFile = fileChooser.showOpenDialog(primaryStage);
                if (selectedFile != null) {
                    Alert alert = AlertComponent.buildAlert();
                    InputStream inputStream = null;
                    try {
                        alert.show();
                        inputStream = new FileInputStream(selectedFile);
                        List<String> strings = FileProcessor.readExcel(inputStream).getFirst();
                        comboBox.getItems().addAll(strings);
                        comboBox.getSelectionModel().select(0);
                        textField.setText(selectedFile.getName());
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    } finally {
                        try {
                            inputStream.close();
                        } catch (IOException ex) {
                            ex.printStackTrace();
                        }
                        if (alert.isShowing()) {
                            alert.close();
                        }
                    }
                }
            });
        } else {
            button.setOnAction(e -> {
                FileChooser fileChooser = new FileChooser();
                fileChooser.setTitle("Choose File");
                fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
                FileChooser.ExtensionFilter excelFilter = new FileChooser.ExtensionFilter("excel文件 (*.xls, *.xlsx)", "*.xls", "*.xlsx");

                fileChooser.getExtensionFilters().add(excelFilter);
                File selectedFile = fileChooser.showOpenDialog(primaryStage);
                if (selectedFile != null) {
                    Alert alert = AlertComponent.buildAlert();
                    InputStream inputStream = null;
                    try {
                        alert.show();
                        inputStream = new FileInputStream(selectedFile);
                        textField.setText(selectedFile.getName());

                        // 处理模板文件，

                    } catch (Exception ex) {
                        ex.printStackTrace();
                    } finally {
                        try {
                            inputStream.close();
                        } catch (IOException ex) {
                            ex.printStackTrace();
                        }
                        if (alert.isShowing()) {
                            alert.close();
                        }
                    }
                }
            });
        }

        final HBox hBoxFile = new HBox(10);
        hBoxFile.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        hBoxFile.getChildren().addAll(label, textField, button);
        return hBoxFile;
    }
}
