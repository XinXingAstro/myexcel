package com.xinxing.excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.text.html.parser.Entity;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.logging.Logger;

import static org.apache.poi.ss.usermodel.CellType.*;

/***
 * 2023-10-23 Merge two excel file
 *
 * Daniel Xin
 */
public class MergeExcel {
    private static Logger logger = Logger.getLogger(MergeExcel.class.getName());

    private Map<String, List<String>> districtMap = new HashMap<>(); // <district, schools>
    private Map<String, List<String>> schoolMap = new HashMap<>(); // <school, teachers>
    private Map<String, String> teacherInfo = new HashMap<>();
    private Set<String> namedSchools = new HashSet<>();

    private void getExcelData() {
        String filePath1 = "D:\\1-团队干部签到表.xlsx";
        String filePath2 = "D:\\2-2023团队干训部学员名单.xlsx";
        Workbook workbook1 = null;
        Workbook workbook2 = null;
        FileInputStream in1 = null;
        FileInputStream in2 = null;
        try {
            File file1 = new File(filePath1);
            File file2 = new File(filePath2);
            if (!file1.exists() || !file2.exists()) {
                logger.warning("指定文件不存在！");
                return;
            }
            in1 = new FileInputStream(file1);
            in2 = new FileInputStream(file2);
            workbook1 = new XSSFWorkbook(in1);
            workbook2 = new XSSFWorkbook(in2);
            Sheet file1Sheet1 = workbook1.getSheetAt(0);
            Sheet file2Sheet1 = workbook2.getSheetAt(0);
            if (file1Sheet1 == null || file2Sheet1 == null) {
                logger.warning("sheet不存在！");
                return;
            }

            // 1. get district-school-teacher map
            String currentDistrict = null;
            String currentSchool = null;
            for (int i = 2; i <= 92; i++) {
                Row row = file1Sheet1.getRow(i);
                assert row != null;
                Cell cell = row.getCell(1);
                CellType cellType = cell.getCellType();
                if (cellType != BLANK) {
                    currentDistrict = cell.getStringCellValue().trim();
                    districtMap.putIfAbsent(currentDistrict, new ArrayList<>());
                }
                cell = row.getCell(2);
                cellType = cell.getCellType();
                if (cellType != BLANK) {
                    currentSchool = cell.getStringCellValue().trim();
                    districtMap.get(currentDistrict).add(currentSchool);
                    schoolMap.putIfAbsent(currentSchool, new ArrayList<>());
                }
                String teacher = row.getCell(3).getStringCellValue().trim();
                schoolMap.get(currentSchool).add(teacher);
            }

            // 2. get list schools
            for (int i = 2; i <= 130; i++) {
                Row row = file2Sheet1.getRow(i);
                assert row != null;
                String school = row.getCell(0).getStringCellValue().trim();
                /*if (namedSchools.contains(school)) {
                    System.out.println(school + " contained!");
                }*/
                namedSchools.add(school);
                if (!schoolMap.containsKey(school)) {
                    schoolMap.putIfAbsent(school, new ArrayList<>());
                }
                String teacher = row.getCell(1).getStringCellValue().trim();
                if (!schoolMap.get(school).contains(teacher)) {
                    schoolMap.get(school).add(teacher);
                }
                String info = row.getCell(2).getStringCellValue().trim();
                teacherInfo.put(teacher, info);
            }

            System.out.println("Named school list length: " + namedSchools.size());

            // 3. generate excel
            String outputFilePath = "D:\\output.xlsx";
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet sheet = outputWorkbook.createSheet();
            int rowIdx = 0;
            // index:0, district:1, schools:2, teachers:3, info:4
            for (String district: districtMap.keySet()) {
                List<String> schoolList = districtMap.get(district);
                boolean districtVisited = false;
                for (String school: schoolList) {
                    if (!namedSchools.contains(school)) continue;
                    List<String> teacherList = schoolMap.get(school);
                    boolean schoolVisited = false;
                    for (String teacher: teacherList) {
                        Row row = sheet.createRow(rowIdx);
                        row.createCell(0).setCellValue(rowIdx);
                        row.createCell(1).setCellValue(districtVisited ? "" : district);
                        districtVisited = true;
                        row.createCell(2).setCellValue(schoolVisited ? "" : school);
                        schoolVisited = true;
                        row.createCell(3).setCellValue(teacher);
                        if (teacherInfo.containsKey(teacher)) {
                            row.createCell(4).setCellValue(teacherInfo.get(teacher));
                        }
                        rowIdx++;
                    }
                    namedSchools.remove(school);
                }
            }
            // append unknown district schools
            for (String school: namedSchools) {
                List<String> teacherList = schoolMap.get(school);
                boolean schoolVisited = false;
                for (String teacher: teacherList) {
                    Row row = sheet.createRow(rowIdx);
                    row.createCell(0).setCellValue(rowIdx);
                    row.createCell(1).setCellValue("未知学区");
                    row.createCell(2).setCellValue(schoolVisited ? "" : school);
                    schoolVisited = true;
                    row.createCell(3).setCellValue(teacher);
                    if (teacherInfo.containsKey(teacher)) {
                        row.createCell(4).setCellValue(teacherInfo.get(teacher));
                    }
                    rowIdx++;
                }
            }
            FileOutputStream out = new FileOutputStream(outputFilePath);
            outputWorkbook.write(out);
            outputWorkbook.close();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (in1 != null) in1.close();
                if (in2 != null) in2.close();
                if (workbook1 != null) workbook1.close();
                if (workbook2 != null) workbook2.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        new MergeExcel().getExcelData();
    }
}
            /* show table1:
            int rowLength = file1Sheet1.getPhysicalNumberOfRows();
            for (int i = 0; i < rowLength; i++) {
                Row row = file1Sheet1.getRow(i);
                if (row == null) {
                    System.out.println("row: " + i + " is null.");
                }
                assert row != null;
                int colLength = row.getPhysicalNumberOfCells();
                for (int j = 0; j < colLength; j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) {
                        System.out.println(i + ":" + j + " NULL ");
                        continue;
                    }
                    CellType cellType = cell.getCellType();
                    if (cellType == NUMERIC) {
                        System.out.print(i + ":" + j + " " + cell.getNumericCellValue());
                    } else if (cellType == STRING) {
                        System.out.print(i + ":" + j + " " + cell.getStringCellValue().trim());
                    } else if (cellType == BLANK) {
                        System.out.print(i + ":" + j + " " + "BLANK");
                    }
                    System.out.print(" ");
                }
                System.out.println();
            }*/