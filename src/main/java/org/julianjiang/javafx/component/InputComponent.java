package org.julianjiang.javafx.component;

import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import org.julianjiang.javafx.Context;

import java.util.function.UnaryOperator;

public class InputComponent {

    Font font;

    public InputComponent(Font font) {
        this.font = font;
    }

    public HBox buildInput(Context context) {
        HBox hBox = new HBox();
        hBox.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        hBox.setSpacing(10); // 设置组件间距


        Label noticeLabel = new Label("注意：");
        String fontTemplate = "-fx-font: 18px \"%s\"; -fx-font-weight: bold; -fx-text-fill: red;";
        noticeLabel.setStyle(fontTemplate);

        Label label1 = new Label();
        label1.setFont(font);
        label1.setText("保留模板前");
        TextField textField = new TextField();

        // 创建一个过滤器，只允许输入整数
        UnaryOperator<TextFormatter.Change> integerFilter = change -> {
            String newText = change.getControlNewText();
            if (newText.matches("-?\\d*")) {
                return change;
            }
            return null;
        };

        // 创建一个TextFormatter，并将过滤器应用于其中
        TextFormatter<String> textFormatter = new TextFormatter<>(integerFilter);
        textField.setTextFormatter(textFormatter);


        textField.setText("0");
        textField.setPrefWidth(150);
        Label label2 = new Label();
        label2.setText("行，并且保留模板后");
        label2.setFont(font);

        TextField textField2 = new TextField();
        textField2.setPrefWidth(150);
        textField2.setText("0");
        TextFormatter<String> textFormatter2 = new TextFormatter<>(integerFilter);
        textField2.setTextFormatter(textFormatter2);

        Label label3 = new Label();
        label3.setText("行！");
        label3.setFont(font);
        hBox.getChildren().addAll(noticeLabel, label1, textField, label2, textField2, label3);

        context.setPreText(textField);
        context.setLastText(textField2);
        return hBox;
    }
}
