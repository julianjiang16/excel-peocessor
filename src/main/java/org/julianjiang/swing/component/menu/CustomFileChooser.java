package org.julianjiang.swing.component.menu;

import javax.swing.*;
import java.awt.*;

public class CustomFileChooser extends JFileChooser {
    public CustomFileChooser() {
        Font font = new Font("Arial", Font.BOLD, 30); // 设置自定义字体和大小
        UIManager.put("FileChooser.font", font); // 设置自定义字体
    }
    // 其他自定义代码
}
