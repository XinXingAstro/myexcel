package com.xinxing.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

public class ExcelReader {
    private static Logger logger = Logger.getLogger(ExcelReader.class.getName());
    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (fileType.equalsIgnoreCase(XLS)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase(XLSX)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    public static List<ExcelDataPO> readExcel(String fileName) {
        Workbook workbook = null;
        FileInputStream inputStream = null;

        try {
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
            File excelFile = new File(fileName);
            if (!excelFile.exists()) {
                logger.warning("指定文件不存在！");
                return null;
            }
            inputStream = new FileInputStream(excelFile);
            workbook = getWorkbook(inputStream, fileType);

            List<ExcelDataPO> resDataList = parseExcel(workbook);

            return resDataList;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        } finally {
            try {
                if (workbook != null) workbook.close();
                if (inputStream != null) inputStream.close();
            } catch (Exception e) {
                e.printStackTrace();
                return null;
            }
        }
    }

    private static List<ExcelDataPO> parseExcel(Workbook workbook) {
        List<ExcelDataPO> resDataList = new ArrayList<ExcelDataPO>();
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = workbook.getSheetAt(sheetNum);
            if (sheet == null) continue;
            int firstRowNum = sheet.getFirstRowNum();
            Row firstRow = sheet.getRow(firstRowNum);
            if (null == firstRow) {
                logger.warning("解析失败，在第一行没有读取到任何数据");
            }

            int rowStart = firstRowNum + 1;
            int rowEnd = sheet.getPhysicalNumberOfRows();
            for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;
                ExcelDataPO rowData = convertRowTodata(row);
                if (rowData == null) {
                    logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
                    continue;
                }
                resDataList.add(rowData);
            }
        }
        return resDataList;
    }

    private static ExcelDataPO convertRowTodata(Row row) {
        ExcelDataPO resData = new ExcelDataPO();

        Cell cell;
        int cellNum = 0;
        cell = row.getCell(cellNum++);
        String name = convertCellValueToString(cell);
        resData.setDevDeploy(5);

        return resData;
    }

    private static String convertCellValueToString(Cell cell) {
        if (cell == null) return null;
        String value = null;
        switch (cell.getCellType()) {
            case NUMERIC:
                Double doubleValue = cell.getNumericCellValue();
                value = new DecimalFormat().format(doubleValue);
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                Boolean booleanValue = cell.getBooleanCellValue();
                value = booleanValue.toString();
                break;
            case BLANK:
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case ERROR:
                break;
            default:
                break;
        }
        return value;
    }
}
