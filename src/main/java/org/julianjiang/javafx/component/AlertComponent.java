package org.julianjiang.javafx.component;

import javafx.scene.control.Alert;

public class AlertComponent {

    public static Alert buildAlert(String title, String msg) {

        // 创建一个警告对话框
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(msg);
        // 禁用关闭按钮和按键
//        alert.getDialogPane().getButtonTypes().clear();
        alert.getDialogPane().getScene().getWindow().setOnCloseRequest(event -> event.consume());
        return alert;
    }
}
