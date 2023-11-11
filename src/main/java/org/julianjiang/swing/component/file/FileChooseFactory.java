package org.julianjiang.swing.component.file;

import javax.swing.*;
import java.awt.*;

public class FileChooseFactory {

    public JPanel buildFileChooserPanel(Font font){

        final JPanel jPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 20, 100));

        final JLabel text = new JLabel();
        text.setText("请选择需要处理的excel文件：");
        jPanel.add(text);

        final JFileChooser jFileChooser = new JFileChooser();
        jFileChooser.setFont(font);

        jPanel.add(jFileChooser);

        return jPanel;
    }
}
