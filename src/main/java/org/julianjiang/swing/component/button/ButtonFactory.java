package org.julianjiang.swing.component.button;

import javax.swing.*;
import java.awt.*;

public class ButtonFactory {


    public JButton buildButton(String name){
        JButton button = new JButton(name);
        button.setPreferredSize(new Dimension(150, 50));
        // 设置按钮2的边距
        button.setMargin(new Insets(10, 30, 10, 30));
        return button;
    }
}
