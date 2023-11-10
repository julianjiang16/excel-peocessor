package org.julianjiang.frame;

import com.google.common.collect.Lists;
import org.julianjiang.component.file.ForgedFileChooseComponent;
import org.julianjiang.component.menu.JComboBoxFactory;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;

public class MainFrame extends JFrame {
    private JButton executeButton;
    private JButton exitButton;

    public MainFrame() {
        Font font = new Font("宋体", Font.BOLD, 31);
        UIManager.put("FileChooser.font", font); // 设置JFileChooser的字体
        // 设置窗口标题
        setTitle("excel加工");
        // 获取屏幕大小
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        // 计算窗口大小为屏幕宽度的70%和高度的60%
        int width = (int) (screenSize.getWidth() * 0.7);
        int height = (int) (screenSize.getHeight() * 0.6);
        // 设置窗口大小
        setSize(width, height);
        // 设置窗口居中对齐
        setLocationRelativeTo(null);
        // 设置关闭窗口时退出程序
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        final ArrayList<String> templates = Lists.newArrayList("模板1", "模板2");
        JComboBoxFactory comboBoxFactory = new JComboBoxFactory(font, width, height);
        final JPanel selectPanel = comboBoxFactory.buildComboBoxPanel("请选择出库模板：", templates);
        // 创建下拉菜单1

        // 创建下拉菜单2
        final ArrayList<String> classTemplates = Lists.newArrayList("---请选择分类字段---", "名称", "日期");
        final JPanel selectPanel2 = comboBoxFactory.buildComboBoxPanel("请选择分类字段：", classTemplates);
        // 创建主面板，并设置布局为垂直盒子布局
        Box mainPanel = Box.createVerticalBox();
        // 设置面板之间的垂直间距
        mainPanel.add(Box.createVerticalStrut(50));

        mainPanel.add(selectPanel);
        mainPanel.add(selectPanel2);

        int fileWidth = (int) (width * 0.7);
        int fileHeight = (int) (height * 0.05);
        final ForgedFileChooseComponent forgedFileChooseComponent = new ForgedFileChooseComponent();
        final JPanel fileJPanel = forgedFileChooseComponent.getFileJPanel(font, fileWidth, fileHeight);
        mainPanel.add(fileJPanel);

        // 创建面板3，放置执行按钮和退出按钮
        JPanel panel3 = new JPanel(new FlowLayout(FlowLayout.CENTER, 100, 10));
        // 创建执行按钮
        executeButton = new JButton("执行");
        // 设置按钮的大小
        executeButton.setPreferredSize(new Dimension(150, 50));
        // 设置按钮的边距
        executeButton.setMargin(new Insets(10, 30, 10, 30));
        // 创建退出按钮
        exitButton = new JButton("退出");

        exitButton.setPreferredSize(new Dimension(150, 50));
        // 设置按钮2的边距
        exitButton.setMargin(new Insets(10, 30, 10, 30));

        executeButton.setFont(font);
        exitButton.setFont(font);
        panel3.add(executeButton);
        panel3.add(exitButton);
        mainPanel.add(panel3);

        // 将主面板添加到窗口中
        add(mainPanel);

        // 显示窗口
        setVisible(true);
    }

    public static void main(String[] args) {
        // 在主线程中创建窗口
        SwingUtilities.invokeLater(MainFrame::new);
    }
}
