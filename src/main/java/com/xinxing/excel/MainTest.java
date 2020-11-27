package com.xinxing.excel;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

public class MainTest {
    private static Logger logger = Logger.getLogger(MainTest.class.getName());

    public static void main(String[] args) {
        // 读取Excel
        String fileName = "demo.xlsx";
        List<ExcelDataPO> res = ExcelReader.readExcel(fileName);

        // 写入Excel
        // 创建数据
        List<ExcelDataPO> dataPOList = new ArrayList<ExcelDataPO>();
        ExcelDataPO dataPO = new ExcelDataPO();
        dataPO.setDevDeploy(5);
        dataPO.setProjectNumber(".....");
        dataPO.setTestDeploy(5);
        dataPOList.add(dataPO);

        // 写入数据到工作簿对象内
        Workbook workbook = ExcelWriter.exportData(dataPOList);

        // 以文件的形式输出工作簿对象
        FileOutputStream fileOut = null;
        try {
            String exportFilePath = "/Users/Dreamer-1/Desktop/myBlog/java解析Excel/writeExample.xlsx";
            File exportFile = new File(exportFilePath);
            if (!exportFile.exists()) {
                exportFile.createNewFile();
            }

            fileOut = new FileOutputStream(exportFilePath);
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            logger.warning("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                if (null != workbook) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.warning("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }

    }
}
