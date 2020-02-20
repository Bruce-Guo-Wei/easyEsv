package eleme;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EasyWork {
    public static void main(String[] args) {
        String proPath = "E:/Project/poi/src/eleme/easy.properties";

        String filepath = "C:/Users/Administrator.USER-20190928MJ/Desktop/800app数据备份/市场管理/Leads潜在学员/Leads潜在学员.csv";
        String outPath = "E:/output";
        Properties pro = new Properties();

        try {
            InputStream is = new FileInputStream(new File(proPath));
            InputStreamReader br = new InputStreamReader(is, Charset.forName("GBK"));
            pro.load(br);
            String headString = pro.getProperty("easywork.headers");
            easyCsv(filepath, outPath, 1001, 2000, headString.split(","));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void easyCsv(String filePath, String fileOutputPath, int startNum, int endNum, String[] headers) throws Exception {
        File file = new File(filePath);
        InputStream is = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(is, "GBK");
        BufferedReader reader = new BufferedReader(isr);

        OutputStream outputStream = null;
        Workbook workbook = new HSSFWorkbook();
        try {
            Sheet sheet = workbook.createSheet("info");
            //第一行
            Row row = sheet.createRow(0);
            for (int c = 0; c < headers.length; c ++) {
                Cell cell = row.createCell(c);
                cell.setCellValue(headers[c]);
            }

            String line;
            int index = 0;
            int creator = 1; // 第一行开始
            int[] columnIndexs = new int[headers.length];
            while ((line = reader.readLine()) != null) {
                String[] cns = line.split(",");
                //省略头
                if (index == 0) {
                    for (int i = 0; i < headers.length; i++) {
                        int id = Arrays.asList(cns).indexOf(headers[i]);
                        if (id == -1) {
                            System.out.println("字段未匹配");
                        }
                        columnIndexs[i] = id;
                    }
                    index ++;
                    continue;
                }

                if (!(index <= endNum && index >= startNum)) {
                    if (index-1 == endNum) {
                        break;
                    }
                    index ++;
                    continue;
                }
                Row row1 = sheet.createRow(creator);
                for (int c = 0; c < columnIndexs.length; c ++) {
                    Cell cell = row1.createCell(c);
                    cell.setCellValue(convert(cns[columnIndexs[c]]));
                }
                creator++;
                index++;
            }
            outputStream = new FileOutputStream(String.format(fileOutputPath  + "/output_line_%s_to_line_%s.xls", startNum, endNum));
            workbook.write(outputStream);
            System.out.println(String.format("文档输出完成:" + fileOutputPath + "/output_line_%s_to_line_%s.xls", startNum, endNum));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                workbook.close();
            }
            if (outputStream != null) {
                outputStream.close();
            }
            if (reader != null) {
                reader.close();
            }
            if (isr != null) {
                isr.close();
            }
            if (is != null) {
                is.close();
            }
        }
    }

    private static String convert (String data) throws Exception {
        Pattern p = Pattern.compile("^\\d{4}\\/\\d{0,2}\\/\\d{0,2} \\d{0,2}:\\d{0,2}$");
        Matcher matcher = p.matcher(data);
        if (matcher.find()) {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/d H:mm");
            SimpleDateFormat sdfTime = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Date date = sdf.parse(data);
            return sdfTime.format(date);
        }
        return data;
    }
}
