package org.julianjiang.component.menu;


import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.plaf.basic.BasicComboBoxEditor;
import java.awt.*;
import java.util.ArrayList;

public class JComboBoxFactory {

    Font font;
    int windowWidth;
    int windowHeight;

    public JComboBoxFactory(Font font, int windowWidth, int windowHeight) {
        this.font = font;
        this.windowWidth = windowWidth;
        this.windowHeight = windowHeight;
    }

    public JPanel buildComboBoxPanel(String labelName, ArrayList templates) {

        int dropdownMenuWidth = (int) (windowWidth * 0.6);
        int dropdownMenuHeight = (int) (windowHeight * 0.05);
        final JPanel subPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 50, 10));
        subPanel.setBorder(new EmptyBorder(0, 0, 0, 0));
        subPanel.setPreferredSize(new Dimension(dropdownMenuWidth, dropdownMenuHeight));
        // label
        final JLabel jLabel = new JLabel();
        jLabel.setFont(font);
        jLabel.setText(labelName);
        subPanel.add(jLabel);
        // select

        JComboBox dropdownMenu = new JComboBox<>();
        templates.stream().forEach(i -> dropdownMenu.addItem(i));
        dropdownMenu.setSelectedIndex(0);
        dropdownMenu.setPreferredSize(new Dimension(dropdownMenuWidth, dropdownMenuHeight));
        dropdownMenu.setFont(font);
        // 设置下拉菜单的渲染器
        dropdownMenu.setRenderer(new DefaultListCellRenderer() {
            private static final long serialVersionUID = 1L;
            private Border border = new LineBorder(Color.RED); // 设置红色边框

            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                Component component = super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (component instanceof JLabel) {
                    JLabel label = (JLabel) component;
                    label.setHorizontalAlignment(SwingConstants.CENTER); // 居中对齐
                    if (index == -1) { // 如果是下拉列表的标题行
                        label.setBorder(border); // 添加红色边框
                    } else {
                        label.setBorder(null); // 清除边框
                    }
                }
                return component;
            }
        });

        // 设置下拉菜单的编辑器
        dropdownMenu.setEditor(new BasicComboBoxEditor() {
            private static final long serialVersionUID = 1L;
            @Override
            public void setItem(Object item) {
                if (item == null) {
                    super.setItem("");
                } else {
                    super.setItem(item);
                }
            }
        });
        subPanel.add(dropdownMenu);

        return subPanel;
    }
}
