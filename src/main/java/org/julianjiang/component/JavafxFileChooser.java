package org.julianjiang.component;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Screen;
import javafx.stage.Stage;

import java.io.File;

public class JavafxFileChooser extends Application {

    @Override
    public void start(Stage primaryStage) {
        BorderPane root = new BorderPane();
        root.setPadding(new Insets(10));
// 获取屏幕大小
        Screen screen = Screen.getPrimary();
        Rectangle2D bounds = screen.getVisualBounds();

        // 计算百分比大小
        double width = bounds.getWidth() * 0.8; // 宽度为屏幕宽度的80%
        double height = bounds.getHeight() * 0.6; // 高度为屏幕高度的60%
        // 创建下拉选项框
        ComboBox<String> comboBox = new ComboBox<>();
        comboBox.getItems().addAll("Option 1", "Option 2", "Option 3");
        comboBox.getSelectionModel().select(0);
        // 创建按钮

        final Label label = new Label();
        label.setText("请选择需要处理的excel：");
        Button button = new Button("选择文件");
        button.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Choose File");
            fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));


            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
//                selectedFileField.setText(selectedFile.getAbsolutePath());
                System.err.println(selectedFile.getAbsolutePath());
            }
        });
        // 创建垂直布局
        VBox vbox = new VBox(10);
        final HBox hBox = new HBox(10);
        hBox.getChildren().addAll(label, button);
        vbox.getChildren().addAll(comboBox, hBox);

        // 将垂直布局放置在主界面的中心
        root.setCenter(vbox);

        Scene scene = new Scene(root, width, height);
        primaryStage.setScene(scene);
        primaryStage.setTitle("JavaFX Application");
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
