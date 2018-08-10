package main;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * @author Yin Changlei
 * @dateTime 2018/8/10 15:03
 */
public class App extends JFrame implements ActionListener {

    private JPanel jPanel;
    private JButton btnChooseFolder;
    private JTextField txtFolder;
    private JButton btnStart;
    private JLabel lblInfo;

    public App() {
        this.setTitle("App");
        this.setSize(520, 460);
        this.setLocation(400, 200);
        this.setResizable(false);

        jPanel = new JPanel();
        jPanel.setLayout(null);

        btnChooseFolder = new JButton("选择文件夹");
        btnChooseFolder.setBounds(10, 10, 100, 30);
        btnChooseFolder.addActionListener(this);
        jPanel.add(btnChooseFolder);

        txtFolder = new JTextField();
        txtFolder.setBounds(10, 50, 400, 30);
        txtFolder.setEnabled(false);
        jPanel.add(txtFolder);

        btnStart = new JButton("开始合并");
        btnStart.setBounds(10, 100, 100, 30);
        btnStart.setFont(Utils.myFont(14));
        btnStart.setForeground(Color.BLUE);
        btnStart.addActionListener(this);
        jPanel.add(btnStart);

        lblInfo = new JLabel("");
        lblInfo.setBounds(10, 150, 400, 30);
        lblInfo.setForeground(Color.RED);
        lblInfo.setFont(Utils.myFont(14));
        jPanel.add(lblInfo);

        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setContentPane(jPanel);
    }

    private void mergeExcel(List<String> excelList) {
        if (excelList.isEmpty()) lblInfo.setText("没有要合并的Excel文件");
        Workbook finalWb = new XSSFWorkbook(); //2007版本，合并后Excel
        for (int i = 0; i < excelList.size(); i++) {
            try {
                String excel = excelList.get(i);
                FileInputStream fileInputStream = new FileInputStream(excel);
                Workbook wb = null;
                if (excel.toLowerCase().endsWith("xls")) {
                    wb = new HSSFWorkbook(fileInputStream);
                } else if (excel.toLowerCase().endsWith("xlsx")) {
                    wb = new XSSFWorkbook(fileInputStream);
                }
                int sheetNum = wb.getNumberOfSheets();
                for (int j = 0; j < sheetNum; j++) {
                    Sheet sheet = wb.getSheetAt(j);
                    Cell cell = sheet.getRow(i).getCell(j);
                    CellStyle cellStyle = cell.getCellStyle();
                    if (sheet.getLastRowNum() <= 0) continue;
                }
            } catch (Exception ex) {

            }
        }
    }

    /**
     * 选择文件夹
     */
    private void chooseFolder() {
        JFileChooser fileChooser = new JFileChooser("D:\\");
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int returnVal = fileChooser.showOpenDialog(fileChooser);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            String filePath = fileChooser.getSelectedFile().getAbsolutePath();
            txtFolder.setText(filePath);
        }
    }

    /**
     * 获取Excel文件列表
     */
    private List<String> findExcels() {
        String folderPath = txtFolder.getText();
        if (Objects.isNull(folderPath) || "".equals(folderPath.trim())) lblInfo.setText("请先选择文件夹");
        lblInfo.setText("");
        List<String> excelList = new ArrayList<>();
        File folder = new File(folderPath);
        File[] files = folder.listFiles();
        if (files == null || files.length == 0) lblInfo.setText("此文件夹下没有文件");
        for (File file : files) {
            if (file.isFile() && ("xlsx".equalsIgnoreCase(getSuffix(file.getName()))
                    || "xls".equalsIgnoreCase(getSuffix(file.getName())))) {
                excelList.add(file.getPath());
            }
        }

        return excelList;
    }

    /**
     * 获取文件后缀
     */
    private String getSuffix(String fileName) {
        return fileName.substring(fileName.lastIndexOf(".") + 1);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getSource() == btnChooseFolder) chooseFolder();
        else if (e.getSource() == btnStart) findExcels();
    }


    public static void main(String[] args) {
        try {
            BeautyEyeLNFHelper.launchBeautyEyeLNF();
            UIManager.put("RootPane.setupButtonVisible", false);
        } catch (Exception e) {
            e.printStackTrace();
        }
        App app = new App();
        app.setVisible(true);
    }
}
