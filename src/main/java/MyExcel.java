import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

public class MyExcel {
    private static Logger logger = Logger.getLogger(MyExcel.class.getName());
    private Map<String, int[]> map = new HashMap<>();

    private void getDevData() {
        String fileName = "C:\\Users\\JMT\\Desktop\\dev.xlsx";
        Workbook workbook = null;
        FileInputStream in = null;
        try {
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warning("指定文件不存在！");
                return;
            }
            in = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(in);
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                logger.warning("sheet不存在！");
                return;
            }
            int rowLength = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowLength; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String projectId = row.getCell(0).getStringCellValue().trim();
                if (!map.containsKey(projectId)) map.put(projectId, new int[3]);
                int[] array = map.get(projectId);
                array[0] = (int) row.getCell(1).getNumericCellValue();
                array[1] = (int) row.getCell(2).getNumericCellValue();
                System.out.println(projectId + ":" + array[0] + ", " + array[1]);
            }
            workbook.close();
            in.close();
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

    private void getTestData() {
        String fileName = "C:\\Users\\JMT\\Desktop\\test.xlsx";
        Workbook workbook = null;
        FileInputStream in = null;
        try {
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warning("指定文件不存在！");
                return;
            }
            in = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(in);
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                logger.warning("sheet不存在！");
                return;
            }
            int rowLength = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowLength; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String projectId = row.getCell(0).getStringCellValue().trim();
                if (!map.containsKey(projectId)) map.put(projectId, new int[3]);
                int[] array = map.get(projectId);
//                array[0] = (int) row.getCell(1).getNumericCellValue();
//                array[1] = (int) row.getCell(2).getNumericCellValue();
                array[2] = (int) row.getCell(2).getNumericCellValue();
                System.out.println(projectId + ":" + array[2]);
            }
            workbook.close();
            in.close();
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
        String fileName = "C:\\Users\\JMT\\Desktop\\123.xlsx";
        Workbook workbook = null;
        FileInputStream in = null;
        FileOutputStream out = null;

        try {
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warning("指定文件不存在！");
                return;
            }
            in = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(in);
            Sheet sheet = workbook.getSheetAt(1);
            if (sheet == null) {
                logger.warning("sheet不存在！");
                return;
            }
            int rowLength = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowLength; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String projectId = row.getCell(0).getStringCellValue().trim();
                if (!map.containsKey(projectId)) continue;
                int[] array = map.get(projectId);
                row.createCell(6).setCellValue(array[0]);
                row.createCell(7).setCellValue(array[1]);
                row.createCell(8).setCellValue(array[2]);
                System.out.println("插入" + projectId + "数据");
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
                logger.warning("关闭资源时出现错误！" + e.getMessage());
            }
        }
    }



    public static void main(String[] args) {
        MyExcel my = new MyExcel();
        my.getDevData();
        my.getTestData();
        my.writeData();
    }
}
