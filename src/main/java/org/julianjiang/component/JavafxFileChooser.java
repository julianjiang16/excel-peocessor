package org.julianjiang.component;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class JavafxFileChooser extends Application {

    @Override
    public void start(Stage primaryStage) {
        BorderPane root = new BorderPane();
        root.setPadding(new Insets(10));

        // 创建下拉选项框
        ComboBox<String> comboBox = new ComboBox<>();
        comboBox.getItems().addAll("Option 1", "Option 2", "Option 3");

        // 创建按钮
        Button button = new Button("按钮");
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
        vbox.getChildren().addAll(comboBox, button);

        // 将垂直布局放置在主界面的中心
        root.setCenter(vbox);

        Scene scene = new Scene(root, 300, 200);
        primaryStage.setScene(scene);
        primaryStage.setTitle("JavaFX Application");
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
