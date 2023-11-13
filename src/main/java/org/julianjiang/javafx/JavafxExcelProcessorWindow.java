package org.julianjiang.javafx;

import com.google.common.collect.Lists;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import javafx.stage.Screen;
import javafx.stage.Stage;
import org.apache.commons.math3.util.Pair;
import org.julianjiang.javafx.component.*;
import org.julianjiang.javafx.model.ExcelTemplate;

public class JavafxExcelProcessorWindow extends Application {

    @Override
    public void start(Stage primaryStage) {

        Context context = new Context(new ExcelTemplate());
        String stringTemplate = "-fx-font: 14px \"%s\"; -fx-font-weight: bold;";
        Font font = new Font("宋体", 14);

        BorderPane root = new CustomBorderPane();
        root.setPadding(new Insets(5));
        // 获取屏幕大小
        Screen screen = Screen.getPrimary();
        Rectangle2D bounds = screen.getVisualBounds();

        // 计算百分比大小
        double width = bounds.getWidth() * 0.4; // 宽度为屏幕宽度的80%
        double height = bounds.getHeight() * 0.6; // 高度为屏幕高度的60%


        Label typeLabel = new Label("请选择分单条件：");
        ComboBoxFactory comboBoxFactory = new ComboBoxFactory(font, typeLabel, stringTemplate);
        Pair typePair = comboBoxFactory.buildComboBox(Lists.newArrayList());
        HBox typeComboBox = (HBox) typePair.getFirst();

        ListView typeCombo = (ListView) typePair.getSecond();
        // 创建按钮
        Label fileLabel = new Label("请选择明细数据文件：");
        FileChooserFactory fileChooserFactory = new FileChooserFactory(font, fileLabel, stringTemplate);
        HBox hboxFile = fileChooserFactory.buildFileChooser(primaryStage, typeCombo, context, false);

        Label tempFileLabel = new Label("请选择输出模板文件：");
        FileChooserFactory tempFileChooserFactory = new FileChooserFactory(font, tempFileLabel, stringTemplate);
        HBox templateBox = tempFileChooserFactory.buildFileChooser(primaryStage, null, context, false);


        Label outputFileLabel = new Label("请选择输出文件夹：");
        FileChooserFactory outputFileChooserFactory = new FileChooserFactory(font, outputFileLabel, stringTemplate);
        HBox outputFBox = outputFileChooserFactory.buildFileChooser(primaryStage, null, context, true);


        Label radioLabel = new Label("是否需要分类汇总（一级分类）：");
        RadioComponent radioComponent = new RadioComponent(font, radioLabel, stringTemplate);
        HBox radioBox = radioComponent.buildRadio(context);

        InputComponent inputComponent = new InputComponent(font);
        VBox noticeBox = inputComponent.buildInput(context);


        ButtonComponent buttonComponent = new ButtonComponent();
        HBox buttonBox = buttonComponent.buildButtonBox(typeCombo, context);

        // 创建垂直布局
        VBox vbox = new VBox(10);
        // 主屏幕
        vbox.getChildren().addAll(hboxFile, typeComboBox, templateBox, outputFBox, radioBox, noticeBox);

        // 将垂直布局放置在主界面的中心
        root.setCenter(vbox);

        root.setBottom(buttonBox);

        Scene scene = new Scene(root, width, height);
        primaryStage.setScene(scene);
        primaryStage.setTitle("文件处理工具");
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
