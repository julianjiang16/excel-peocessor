package org.julianjiang.component;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class CustomFileChooserExample extends Application {

    @Override
    public void start(Stage primaryStage) {
        TextField selectedFileField = new TextField();
        selectedFileField.setEditable(false);
        Button chooseFileButton = new Button("Choose File");

        chooseFileButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Choose File");
            fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                selectedFileField.setText(selectedFile.getAbsolutePath());
            }
        });

        VBox root = new VBox(10);
        root.getChildren().addAll(selectedFileField, chooseFileButton);

        Scene scene = new Scene(root, 300, 200);
        primaryStage.setScene(scene);
        primaryStage.setTitle("Custom File Chooser");
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
