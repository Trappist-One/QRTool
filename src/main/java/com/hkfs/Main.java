package com.hkfs;

import com.hkfs.utils.ExcelUtil;
import com.hkfs.utils.QRCodeUtil;

import javax.swing.*;
import javax.swing.border.LineBorder;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;

/**
 * @Description:
 * @Author: Bruce Lee
 * @Date: 2021/11/2 11:35
 */
public class Main extends JFrame {

    private static String URL_PREFIX = "";

    private ExecutorService service = Executors.newCachedThreadPool(new ThreadFactory() {
        @Override
        public Thread newThread(Runnable r) {
            return new Thread(r, "output");
        }
    });


    public static void main(String[] args) throws Exception {
        Main main = new Main();
        main.initWindow();

    }

    private void initWindow() {
        this.setTitle("生成二维码");
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setSize(450, 300);
        this.setLocationRelativeTo(null); // 居中
        this.setLayout(null);

        JLabel excelPathLabel = new JLabel("文件路径：");
        excelPathLabel.setBounds(new Rectangle(20, 20, 70, 25));
        this.add(excelPathLabel);
        JTextField excelPathField = new JTextField(30);
        excelPathField.setBounds(new Rectangle(100, 20, 300, 25));
        this.add(excelPathField);

        JLabel savePathLabel = new JLabel("保存路径：");
        savePathLabel.setBounds(new Rectangle(20, 50, 70, 25));
        this.add(savePathLabel);
        JTextField savePathField = new JTextField(30);
        savePathField.setBounds(new Rectangle(100, 50, 300, 25));
        this.add(savePathField);

        JLabel urlPathLabel = new JLabel("链接前缀：");
        urlPathLabel.setBounds(new Rectangle(20, 80, 70, 25));
        this.add(urlPathLabel);
        JTextField urlPathField = new JTextField(30);
        urlPathField.setBounds(new Rectangle(100, 80, 300, 25));
        urlPathField.setText("");
        this.add(urlPathField);

        JButton button = new JButton("生成");
        button.setPreferredSize(new Dimension(20, 30));
        button.setBounds(new Rectangle(20, 110, 70, 20));
        this.add(button);

        JScrollPane jScrollPane = new JScrollPane();
        JTextArea jTextArea = new JTextArea();
        jScrollPane.setBounds(new Rectangle(100, 140, 300, 100));
        jTextArea.setBorder(new LineBorder(Color.GRAY, 1));
        jTextArea.setBackground(new Color(200, 200, 200));
        jTextArea.setForeground(Color.BLUE);
        jTextArea.setEditable(false);
        jScrollPane.setViewportView(jTextArea);
        this.add(jScrollPane);
        this.setVisible(true);

        button.addActionListener(e -> {
            service.submit(()->{
                button.setEnabled(false);
                String excelPathFieldText = excelPathField.getText();
                String savePathFieldText = savePathField.getText();
                URL_PREFIX = urlPathField.getText();
                String result = null;
                if (excelPathFieldText == null || "".equals(excelPathFieldText)) {
                    result = "\n" + "文件路径为空！";
                }

                if (savePathFieldText == null || "".equals(excelPathFieldText)) {
                    result = "\n" + "保存路径为空！";
                }

                if (URL_PREFIX == null || "".equals(excelPathFieldText)) {
                    result = "\n" + "链接前缀为空！";
                }

                if (result == null) {
                    result = createQRCode(excelPathFieldText, savePathFieldText, jTextArea);
                }

                if (result == null) {
                    jTextArea.append("\n" +"完成！");
                } else {
                    jTextArea.append(result);
                }
                button.setEnabled(true);
            });
        });


    }

    private String createQRCode(String filePath, String savePath, JTextArea jTextArea) {
        String result = null;
        File f = new File(filePath);
        List<Map<String, Object>> list = null;
        try {
            jTextArea.setText("解析文件开始...");
            list = ExcelUtil.readeExcelData(new FileInputStream(f), 0, 0, 1);
            jTextArea.append("\n" + "解析文件完成.");
        } catch (Exception e) {
            result = "\n解析文件失败！";
        }
        try {
            jTextArea.append("\n" + "生成二维码开始...");
            int size = list.size();
            for (int i = 0; i < size; i++) {
                Map<String, Object> map = list.get(i);
                String agentNo = (String) map.get("代理人编码");
                String name = (String) map.get("姓名");
                String text = URL_PREFIX + agentNo;
                QRCodeUtil.encode(text, null, savePath, agentNo + "_" + name, true);
                jTextArea.append("\n已完成" + (i + 1) +"/" + size);

            }
        } catch (Exception e) {
            result = "\n生成二维码失败！";
        }

        return result;
    }
}
