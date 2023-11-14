package org.julianjiang.javafx.component;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.SelectionMode;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.math3.util.Pair;

import java.lang.reflect.InvocationTargetException;
import java.util.List;

public class ComboBoxFactory {

    Font font;
    String fontTemplate;
    Label label;

    public ComboBoxFactory(Font font, Label label,String fontTemplate) {
        this.font = font;
        this.fontTemplate = String.format(fontTemplate, font.getName());
        this.label = label;
        this.label.setStyle(fontTemplate);
        this.label.setPrefWidth(300);
    }

    public void updateLabelName(String labelName) {
        this.label.setText(labelName);
    }

    public Pair buildComboBox(List<String> template) {
        HBox hBox = new HBox();
        hBox.setPadding(new Insets(30, 0, 0, 20)); // 设置左边距
        hBox.setSpacing(10); // 设置组件间距

        ListView<String> listView = new ListView<>();
        listView.setPrefWidth(300);
        listView.setPrefHeight(150);
        listView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        ObservableList<String> items = FXCollections.observableArrayList();
        listView.setItems(items);
        Label label = new Label();
        try {
            BeanUtils.copyProperties(label, this.label);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        hBox.getChildren().addAll(label, listView);
        return Pair.create(hBox,listView);
    }
}
