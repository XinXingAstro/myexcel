package com.xinxing.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SplitRow {
    private static Logger logger = Logger.getLogger(MyExcel.class.getName());

    //excel里有的cell有多行内容，每行都是一个项目，多个项目右边一格对应这些项目的上线日期
    //现在要求每个项目一行右边对应该项目的上线日期，项目只要项目编号即可。
    private void splitRow() {
        String sourceFile = "C:\\Users\\JMT\\Desktop\\1.xlsx";
        String targetFile = "C:\\Users\\JMT\\Desktop\\2.xlsx";
        Workbook workbook = null;
        FileInputStream in = null;
        FileOutputStream out = null;

        try {
            File excelFile = new File(sourceFile);
            if (!excelFile.exists()) {
                logger.warning("指定文件不存在！");
                return;
            }
            in = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(in);
            Sheet sheet0 = workbook.getSheetAt(0);
            Sheet sheet1 = workbook.getSheetAt(1);
            if (sheet0 == null) {
                logger.warning("sheet不存在！");
                return;
            }
            int idx = 0;
            Pattern pat = Pattern.compile("\\d{4}-\\d{3}");//直接匹配项目编号
            for (int r = 2; r < 113; r++) {
                Row row0 = sheet0.getRow(r);
                if (row0 == null) continue;
                String raw = row0.getCell(4).getStringCellValue().trim();
                Date key = row0.getCell(5).getDateCellValue();
                Matcher mat = pat.matcher(raw);
                while (mat.find()) {
                    Row row1 = sheet1.createRow(idx++);
                    row1.createCell(0).setCellValue(mat.group());
                    row1.createCell(1).setCellValue(key);
//                    System.out.println(mat.group());
                }
//                for (String project : projects.split("\n")) {
//                    Row row1 = sheet1.createRow(idx++);
//                    row1.createCell(0).setCellValue(project);
//                    row1.createCell(1).setCellValue(key);
//                    System.out.println("插入" + project + "：" + key);
//                }
            }
            in.close();
            File outFile = new File(targetFile);
            out = new FileOutputStream(outFile);
            workbook.write(out);
            workbook.close();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (in != null) in.close();
                if (workbook != null) workbook.close();
                if (out != null) out.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        SplitRow sr = new SplitRow();
        sr.splitRow();
    }
}
