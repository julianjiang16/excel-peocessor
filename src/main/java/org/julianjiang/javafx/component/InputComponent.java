package org.julianjiang.javafx.component;

import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.text.Font;
import org.julianjiang.javafx.Context;

import java.util.function.UnaryOperator;

public class InputComponent {

    Font font;

    public InputComponent(Font font) {
        this.font = font;
    }

    public VBox buildInput(Context context) {
        final VBox vBox = new VBox();

        HBox noticeHBox = new HBox();
        noticeHBox.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        noticeHBox.setSpacing(10); // 设置组件间距

        HBox titleHBox = new HBox();
        titleHBox.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        titleHBox.setSpacing(10); // 设置组件间距


        Label noticeLabel = new Label("注意*：");
        String fontTemplate = "-fx-font: 14px \"%s\"; -fx-font-weight: bold; -fx-text-fill: red;";
        noticeLabel.setStyle(fontTemplate);

        Label label1 = new Label();
        label1.setFont(font);
        label1.setText("选择的模板文件中，前");
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
        textField.setPrefWidth(60);
        textField.setFont(font);
        Label label2 = new Label();
        label2.setText("行为页头，后");
        label2.setFont(font);

        TextField textField2 = new TextField();
        textField2.setPrefWidth(60);
        textField2.setFont(font);
        textField2.setText("0");
        TextFormatter<String> textFormatter2 = new TextFormatter<>(integerFilter);
        textField2.setTextFormatter(textFormatter2);

        Label label3 = new Label();
        label3.setText("行为页尾（生成文件是会保留页头页尾，必填！！！）；");
        label3.setFont(font);
        noticeHBox.getChildren().addAll(noticeLabel, label1, textField, label2, textField2, label3);

        context.getExcelTemplate().setPreText(textField);
        context.getExcelTemplate().setLastText(textField2);

        // 选择的模板文件中，列名处于第 N 行。  分类汇总行（样式）处于第 N 行；

        final Label titleLabel = new Label("列名处于第");
        titleLabel.setFont(font);

        final TextField titleTextField = new TextField("0");
        titleTextField.setFont(font);

        TextFormatter<String> titleTextFormatter = new TextFormatter<>(integerFilter);
        titleTextField.setPrefWidth(60);
        titleTextField.setTextFormatter(titleTextFormatter);

        final Label titleRowLabel = new Label("行(必填)，分类汇总行（样式）处于第");
        titleRowLabel.setFont(font);

        final TextField typeTextField = new TextField("0");
        typeTextField.setFont(font);
        typeTextField.setPrefWidth(60);
        TextFormatter<String> typeTextFormatter = new TextFormatter<>(integerFilter);
        typeTextField.setTextFormatter(typeTextFormatter);

        final Label typeRowLabel = new Label("行(非必填)，明细数据行（样式）处于第");
        typeRowLabel.setFont(font);


        final TextField detailTextField = new TextField("0");
        detailTextField.setFont(font);
        detailTextField.setPrefWidth(60);
        TextFormatter<String> detailTextFormatter = new TextFormatter<>(integerFilter);
        detailTextField.setTextFormatter(detailTextFormatter);

        final Label detailRowLabel = new Label("行(必填).");
        detailRowLabel.setFont(font);


        context.getExcelTemplate().setTitleText(titleTextField);
        context.getExcelTemplate().setTypeText(typeTextField);
        context.getExcelTemplate().setDetailText(detailTextField);

        titleHBox.getChildren().addAll(titleLabel, titleTextField, titleRowLabel, typeTextField, typeRowLabel, detailTextField, detailRowLabel);
        vBox.getChildren().addAll(noticeHBox, titleHBox);
        return vBox;
    }
}
