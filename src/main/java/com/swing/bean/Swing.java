package com.swing.bean;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

import com.alibaba.fastjson.JSONObject;

import org.apache.poi.hssf.usermodel.*;
import org.apache.http.entity.StringEntity;
import org.apache.poi.ss.format.CellDateFormatter;

public class Swing extends JFrame {

    private static final long serialVersionUID = 1L;
    /* 主窗体里面的若干元素 */
    private JFrame mainForm = new JFrame("excel文件解析");
    private JLabel label1 = new JLabel("请选择需要进行处理的excel文件：");
    private JLabel label2 = new JLabel("请选择处理后的图片存储在哪里：");
    public static JTextField sourcefile = new JTextField(); // 选择待加密或解密文件路径的文本域
    public static JTextField targetfile = new JTextField(); // 选择加密或解密后文件路径的文本域
    public static JButton buttonBrowseSource = new JButton("浏览"); // 浏览按钮
    public static JButton buttonBrowseTarget = new JButton("浏览"); // 浏览按钮
    public static JButton buttonEncrypt = new JButton("确定"); // 加密按钮

    public Swing() {
        Container container = mainForm.getContentPane();
        /* 设置主窗体属性 */
        mainForm.setSize(400, 300);// 设置主窗体大小
        mainForm.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);// 设置主窗体关闭按钮样式
        mainForm.setLocationRelativeTo(null);// 设置居于屏幕中央
        mainForm.setResizable(false);// 设置窗口不可缩放
        mainForm.setLayout(null);
        mainForm.setVisible(true);// 显示窗口
        /* 设置各元素位置布局 */
        label1.setBounds(30, 10, 300, 30);
        sourcefile.setBounds(50, 50, 200, 30);
        buttonBrowseSource.setBounds(270, 50, 60, 30);
        label2.setBounds(30, 90, 300, 30);
        targetfile.setBounds(50, 130, 200, 30);
        buttonBrowseTarget.setBounds(270, 130, 60, 30);
        buttonEncrypt.setBounds(100, 180, 60, 30);

        buttonBrowseSource.addActionListener(new BrowseAction());
        buttonBrowseTarget.addActionListener(new BrowseAction());
        buttonEncrypt.addActionListener(new ManageExcel());

        sourcefile.setEditable(false);// 设置源文件文本域不可手动修改
        targetfile.setEditable(false);// 设置目标位置文本域不可手动修改
        container.add(label1);
        container.add(label2);
        container.add(sourcefile);
        container.add(targetfile);
        container.add(buttonBrowseSource);
        container.add(buttonBrowseTarget);
        container.add(buttonEncrypt);
    }

    public class BrowseAction implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            if (e.getSource().equals(Swing.buttonBrowseSource)) {
                JFileChooser fcDlg = new JFileChooser();
                fcDlg.setDialogTitle("请选择需要进行处理的excel文件...");
                FileNameExtensionFilter filter = new FileNameExtensionFilter(
                        "文本文件(*.xls)", "xls");
                fcDlg.setFileFilter(filter);
                int returnVal = fcDlg.showOpenDialog(null);
                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    String filepath = fcDlg.getSelectedFile().getPath();
                    Swing.sourcefile.setText(filepath);
                }
            } else if (e.getSource().equals(Swing.buttonBrowseTarget)) {
                JFileChooser fcDlg = new JFileChooser();
                fcDlg.setDialogTitle("请选择加密或解密后的文件存放目录");
                fcDlg.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int returnVal = fcDlg.showOpenDialog(null);
                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    String filepath = fcDlg.getSelectedFile().getPath();
                    Swing.targetfile.setText(filepath);
                }
            }
        }
    }

    public class ManageExcel implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            if (Swing.sourcefile.getText().isEmpty()) {
                JOptionPane.showMessageDialog(null, "请选择待解析的excel文件！");
            } else if (Swing.targetfile.getText().isEmpty()) {
                JOptionPane.showMessageDialog(null, "请选择解析后文件存放目录！");
            } else {
                String sourcepath = Swing.sourcefile.getText();
                String targetpath = Swing.targetfile.getText();
                File file = new File(sourcepath);
                String filename = file.getName();
                File dir = new File(targetpath);
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                if (file.exists() && dir.isDirectory()) {
                    OutputStream out = null;
                    HSSFSheet sheetPicture;
                    File fileOut = new File(targetpath + File.separator + "demo" + File.separator + "picture.xls");
                    if (!fileOut.getParentFile().exists()) { //如果文件的目录不存在
                        fileOut.getParentFile().mkdirs(); //创建目录
                    }
                    try (OutputStream fileOutputStream = new FileOutputStream(fileOut);
                         HSSFWorkbook wb = new HSSFWorkbook();
                         HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(sourcepath)))) {
                        short index = 0;
                        short indexCol = 19;
                        sheetPicture = wb.createSheet("图片");
                        HSSFSheet sheet;
                        Map<String, Double> LowerAndUpperMap = new HashMap<>();

                        if (workbook.getNumberOfSheets() < 2) {
                            JOptionPane.showMessageDialog(null, "只有一个sheet表，没有重新解析的必要");
                            return;
                        }

                        List<Data> dataList = new ArrayList<>();
                        String deviceName = String.valueOf(workbook.getSheetAt(0).getRow(0).getCell(0)).split("-")[1];
                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            if (workbook.getSheetName(i).equals("数据报告")) {
                                sheet = workbook.getSheetAt(i);
                                for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {
                                    HSSFRow row = sheet.getRow(j);
                                    if (row != null) {
                                        for (int k = 0; k < row.getLastCellNum(); k++) {
                                            String str = String.valueOf(row.getCell(k));
                                            if ((str.contains("报警上限"))) {
                                                String strData = String.valueOf(row.getCell(k + 1));
                                                if ((strData.contains("未设置"))) {
                                                    LowerAndUpperMap.put(str, Double.MAX_VALUE);
                                                } else {
                                                    strData = strData.split("%")[0].split("°C")[0];
                                                    LowerAndUpperMap.put(str, Double.valueOf(strData));
                                                }
                                            } else if ((str.contains("报警下限"))) {
                                                String strData = String.valueOf(row.getCell(k + 1));
                                                if ((strData.contains("未设置"))) {
                                                    LowerAndUpperMap.put(str, Double.MIN_VALUE);
                                                } else {
                                                    strData = strData.split("%")[0].split("°C")[0];
                                                    LowerAndUpperMap.put(str, Double.valueOf(strData));
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (workbook.getSheetName(i).equals("历史数据")) {
                                sheet = workbook.getSheetAt(i);
                                HSSFRow row = sheet.getRow(1);
                                int length = row.getLastCellNum();
                                for (int j = 1; j < length - 1; j++) {
                                    Data data = new Data();
                                    List<List<Object>> dataPointList = new ArrayList<>();
                                    String str = String.valueOf(sheet.getRow(1).getCell(j));
                                    str = str.split("\\(")[0];
                                    data.setName(str);
                                    data.setLowLimit(LowerAndUpperMap.get(str + "报警下限"));
                                    data.setUpperLimit(LowerAndUpperMap.get(str + "报警上限"));
                                    for (int k = 2; k < sheet.getLastRowNum() + 1; k++) {
                                        List<Object> doubles = new ArrayList<>();
                                        Date dt1;
                                        try{
                                            String value = String.valueOf(sheet.getRow(k).getCell(length - 1));
                                            dt1 = sdf.parse(value);
                                        }catch (Exception e1){
                                           dt1 = sheet.getRow(k).getCell(length - 1).getDateCellValue();
                                        }
                                        long ts1 = dt1.getTime();
                                        doubles.add(0, ts1);
                                        doubles.add(1, Double.valueOf(String.valueOf(sheet.getRow(k).getCell(j))));
                                        if (Double.valueOf(String.valueOf(sheet.getRow(k).getCell(j))) > data.getMax()) {
                                            data.setMax(Double.valueOf(String.valueOf(sheet.getRow(k).getCell(j))));
                                        }
                                        if (Double.valueOf(String.valueOf(sheet.getRow(k).getCell(j))) < data.getMin()) {
                                            data.setMin(Double.valueOf(String.valueOf(sheet.getRow(k).getCell(j))));
                                        }
                                        dataPointList.add(doubles);
                                    }
                                    data.setValue(dataPointList);
                                    dataList.add(data);
                                }
                            }
                        }

                        for (int x = 0, len = dataList.size(); x < len; x++) {
                            String data = JSONObject.toJSONString(dataList.get(x).getValue());
                            byte[] bytes = null;
                            boolean up_initial, low_initial;
                            Double upLimit, lowLimit;
                            if (dataList.get(x).getUpperLimit().equals(Double.MAX_VALUE)) {
                                up_initial = false;
                                upLimit = 0D;
                            } else {
                                up_initial = true;
                                upLimit = dataList.get(x).getUpperLimit();
                            }
                            if (dataList.get(x).getLowLimit().equals(Double.MIN_VALUE)) {
                                low_initial = false;
                                lowLimit = 0D;
                            } else {
                                low_initial = true;
                                lowLimit = dataList.get(x).getLowLimit();
                            }

                            if (dataList.get(x).getName().contains("温度")) {
                                bytes = getImageByData(deviceName, data, dataList.get(x).getName(), "℃", lowLimit, upLimit, low_initial, up_initial, dataList.get(x).getMin(), dataList.get(x).getMax());
                            } else if (dataList.get(x).getName().contains("湿度")) {
                                bytes = getImageByData(deviceName, data, dataList.get(x).getName(), "%", lowLimit, upLimit, low_initial, up_initial, dataList.get(x).getMin(), dataList.get(x).getMax());
                            } else {
                                JOptionPane.showMessageDialog(null, "不是应有的解析类型");
                                return;
                            }
                            HSSFPatriarch patriarchTemperature = sheetPicture.createDrawingPatriarch();
                            HSSFClientAnchor anchorTemperature = new HSSFClientAnchor(0, 0, 500, 255, (short) 1, index, (short) 10, indexCol);
                            patriarchTemperature.createPicture(anchorTemperature, wb.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG));
                            index += 19;
                            indexCol += 19;
                        }
                        wb.write(fileOutputStream);
                        JOptionPane.showMessageDialog(null, "解析excel成功！");
                    } catch (Exception x) {
                        System.out.println("there are some err in the procedure");
                    }
                } else if (!file.exists()) {
                    JOptionPane.showMessageDialog(null, "待解析文件不存在！");
                } else {
                    JOptionPane.showMessageDialog(null, "解析后图片存放目录不存在！");
                }
            }
        }
    }

    public byte[] getImageByData(String device_name, String data,
                                 String title, String unit, Double lower, Double upper, boolean low_initial, boolean up_initial, Double min,
                                 Double max) {
        String url = "http://127.0.0.1:3003/";
        String low = low_initial == false ? "" : "{color: \'#ffdc00\',width: 2, color: 'green', dashStyle: 'shortdash',value: "
                + lower
                + ",zIndex:1}";
        String up = up_initial == false ? "" : ",{color: \'#ff851b\',width: 2, color: 'red', dashStyle: 'shortdash',value: "
                + upper
                + ",zIndex:1}";
        String infile = "{\"infile\":\"{title: {text: \'"
                + device_name
                + title
                + "曲线图\'},credits : {enabled : false},xAxis: {type: \'datetime\' ,gridLineColor: \'#197F07\',gridLineWidth: 1 },yAxis: {gridLineWidth: 1,gridLineDashStyle: \'longdash\',gridLineWidth: 1,max:"
                + (max + 10)
                + ",min:"
                + (min - 10)
                + ",plotLines:["
                + low
                + up
                + "],title: {text: \'"
                + title
                + "("
                + unit
                + ")\'}},series:[{name: \'"
                + device_name
                + title
                + "曲线图\',data:"
                + data
                + ",tooltip: {valueDecimals: 1,valueSuffix: \'s\'}}]}\",\"constr\":\"StockChart\",\"globaloptions\":\"{global:{useUTC: false}} \"}";
        StringEntity stringEntity;
        byte[] bytes = null;
        stringEntity = new StringEntity(infile, "UTF-8");
        String res = HttpUtil.Post(url, stringEntity, null, true, 1);
        Base64Util base64Util = new Base64Util();
        bytes = base64Util.base64ToImage(res);
        return bytes;
    }

    public static void startPhantomjsServer() {
        final Process process;
        ArrayList<String> commands = new ArrayList<>();
        String os_name = System.getProperty("os.name");
        String path;

        if (os_name.contains("linux") || os_name.contains("Linux")
                || os_name.contains("LINUX")) {
            commands.add("phantomjs");
        } else if (os_name.contains("windows") || os_name.contains("Windows")) {
            commands.add("./res/IDEA-workspace/" + "phantomjs-1.9.7-windows/"
                    + "phantomjs.exe");
        }
        commands.add("--output-encoding=" + "utf8");

        path = "./res/IDEA-workspace/" + "phantomjs/";
        commands.add(path + "highcharts-convert.js");
        commands.add("-host");
        commands.add("127.0.0.1");
        commands.add("-port");
        commands.add("" + "3003");

        try {
            process = new ProcessBuilder(commands).start();
            final BufferedReader bufferedReader = new BufferedReader(
                    new InputStreamReader(process.getInputStream()));
            String readLine = bufferedReader.readLine();
            if (readLine == null || !readLine.contains("ready")) {
                process.destroy();
                throw new RuntimeException("Error, PhantomJS couldnot start");
            }
            Runtime.getRuntime().addShutdownHook(new Thread() {
                @Override
                public void run() {
                    if (process != null) {
                        try {
                            process.getErrorStream().close();
                            process.getInputStream().close();
                            process.getOutputStream().close();
                        } catch (IOException e) {
                        }
                        process.destroy();
                    }
                }
            });
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        startPhantomjsServer();
        new Swing();
    }
}
