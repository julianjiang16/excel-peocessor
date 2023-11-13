package org.julianjiang.javafx.component;

import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import org.apache.commons.lang.StringUtils;
import org.julianjiang.javafx.Context;

public class RadioComponent {
    Font font;
    String fontTemplate;
    Label label;

    public RadioComponent(Font font, Label label, String fontTemplate) {
        this.font = font;
        this.fontTemplate = String.format(fontTemplate, font.getName());
        this.label = label;
        this.label.setStyle(fontTemplate);
        this.label.setPrefWidth(300);
    }

    public HBox buildRadio(Context context) {

        HBox hBox = new HBox();
        hBox.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        hBox.setSpacing(10); // 设置组件间距
        // 创建单选按钮
        RadioButton falseButton = new RadioButton("否");
        falseButton.setFont(font);

        RadioButton trueButton = new RadioButton("是");
        trueButton.setFont(font);
        trueButton.setPadding(new Insets(0, 0, 0, 100));

        // 创建一个ToggleGroup，并将单选按钮添加到ToggleGroup中
        ToggleGroup toggleGroup = new ToggleGroup();
        falseButton.setToggleGroup(toggleGroup);
        trueButton.setToggleGroup(toggleGroup);

        toggleGroup.selectedToggleProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue != null && StringUtils.isNotBlank(newValue.toString())) {
                if (newValue.toString().indexOf("是") > -1) {
                    context.setTypeFlag(true);
                } else {
                    context.setTypeFlag(false);
                }
            }
        });

        // 设置默认选中的单选按钮
        falseButton.setSelected(true);
        context.setTypeFlag(false);

        hBox.getChildren().addAll(label, falseButton, trueButton);
        return hBox;
    }
}
