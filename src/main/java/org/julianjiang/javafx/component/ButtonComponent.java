package org.julianjiang.javafx.component;

import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.MultipleSelectionModel;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;

public class ButtonComponent {


    public HBox buildButtonBox(ListView listView) {
        HBox hBox = new HBox();
        hBox.setAlignment(Pos.CENTER);
        hBox.setPadding(new Insets(10, 0, 30, 20)); // 设置左边距
        hBox.setSpacing(100); // 设置组件间距
        Font f = Font.font("Arial", 30);
        Button process = new Button();
        process.setText("执行");
        process.setFont(f);

        process.setOnAction(actionEvent -> {
            MultipleSelectionModel<String> selectionModel = listView.getSelectionModel();
            ObservableList<String> selectedItems = selectionModel.getSelectedItems();
            for (String item : selectedItems) {
                System.out.println(item);
            }
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
