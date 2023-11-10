package org.julianjiang.component.file;

import javax.swing.*;
import java.awt.*;
import java.io.File;

public class ForgedFileChooseComponent {


    public JPanel getFileJPanel(Font font, int fileWidth, int fileHeight) {
        final JPanel jPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 20, 50));

        // label
        final JLabel jLabel = new JLabel();
        jLabel.setText("请选择需要处理的excel文件：");
        jLabel.setFont(font);
        // text
        final JTextField jTextField = new JTextField();

        jTextField.setPreferredSize(new Dimension(fileWidth, fileHeight));
        jTextField.setText("");
        jTextField.setEditable(false);
        jTextField.setFont(font);
        // button

        final JButton button = new JButton();
        button.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFont(font);
            int height = (int) (fileWidth * 0.5);
            fileChooser.setPreferredSize(new Dimension(fileWidth, height));
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                String fileName = selectedFile.getName();
                jTextField.setText(fileName);
            }
        });
        button.setText("请选择...");
        button.setFont(font);

        jPanel.add(jLabel);
        jPanel.add(jTextField);
        jPanel.add(button);
        return jPanel;
    }
}
