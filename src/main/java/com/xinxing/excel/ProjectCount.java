package com.xinxing.excel;

import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ProjectCount {
    //<project_id, [清单个数，开发上线次数，测试上线次数]>
    private Map<String, int[]> map;

    public ProjectCount() {
        map = new HashMap<>();
    }
    
    private void getData() {
        String fileName = "C:\\Users\\JMT\\Desktop\\软件质量量化指标统计表 (202101-202107).xls";
        Workbook workbook = null;
        FileInputStream in = null;
        try {
            File excelFile = new File(fileName);
            in = new FileInputStream(excelFile);
            workbook = new HSSFWorkbook(in);
            System.out.println("文件装入完成");
            //填入开发数据
            Sheet sheet1 = workbook.getSheet("Sheet1-开发");
            if (sheet1 == null) {
                System.out.println("Sheet1-开发不存在！");
            }
            int rowLength = sheet1.getPhysicalNumberOfRows();
            for (int i = 1; i < rowLength; i++) {
                Row row = sheet1.getRow(i);
                if (row == null) continue;
                String projectId = row.getCell(0).getStringCellValue().trim();
                if (!map.containsKey(projectId)) map.put(projectId, new int[3]);
                int[] array = map.get(projectId);
                //填入开发清单个数
                array[0] = (int) row.getCell(1).getNumericCellValue();
                //填入开发上线次数
                array[1] = (int) row.getCell(2).getNumericCellValue();
                System.out.println(projectId + " 清单个数: " + array[0] + "; 开发上线次数: " + array[1]);
            }
            //填入测试数据
            Sheet sheet2 = workbook.getSheet("Sheet3-测试");
            if (sheet2 == null) {
                System.out.println("Sheet3-测试不存在！");
            }
            rowLength = sheet2.getPhysicalNumberOfRows();
            for (int i = 1; i < rowLength; i++) {
                Row row = sheet2.getRow(i);
                if (row == null) continue;
                String projectId = row.getCell(0).getStringCellValue().trim();
                if (!map.containsKey(projectId)) map.put(projectId, new int[3]);
                int[] array = map.get(projectId);
                //填入测试上线次数
                array[2] = (int) row.getCell(2).getNumericCellValue();
                System.out.println(projectId + " 测试上线次数: " + array[2]);
            }
            
            in.close();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (in != null) in.close();
                if (workbook != null) workbook.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private void writeData() {
        String fileName = "C:\\Users\\JMT\\Desktop\\软件质量量化指标统计表 (202101-202107).xls";
        Workbook workbook = null;
        FileInputStream in = null;
        FileOutputStream out = null;

        try {
            File excelFile = new File(fileName);
            in = new FileInputStream(excelFile);
            workbook = new HSSFWorkbook(in);
            Sheet sheet = workbook.getSheet("Sheet2");
            if (sheet == null) System.out.println("Sheet2不存在！");
            int rowLength = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowLength; i++) {
                Row row = sheet.getRow(i);
                if (row == null) break;
                Cell pidCell = row.getCell(1);
                if (pidCell == null) break;
                String pid = pidCell.getStringCellValue().trim();
                if (pid == null) break;
                if (!map.containsKey(pid)) {
                    row.createCell(23).setCellValue(0);
                    row.createCell(24).setCellValue(0);
                    row.createCell(25).setCellValue(0);
                } else {
                    int[] data = map.get(pid);
                    row.createCell(23).setCellValue(data[0]);
                    row.createCell(24).setCellValue(data[1]);
                    row.createCell(25).setCellValue(data[2]);
                    System.out.println(pid+" 上线清单个数："+data[0]+" 开发上线次数: "+data[1]+" 修改上线次数: "+data[2]);
                }
            }
            in.close();
            out = new FileOutputStream(fileName);
            workbook.write(out);
            workbook.close();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) workbook.close();
                if (in != null) in.close();
                if (out != null) out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        ProjectCount pc = new ProjectCount();
        pc.getData();
        pc.writeData();
    }
}
