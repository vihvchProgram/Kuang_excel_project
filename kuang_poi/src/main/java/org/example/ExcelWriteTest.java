package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

public class ExcelWriteTest {

    String PATH = "..\\..\\";

    @Test
    public void testWrite03() throws Exception {
        // 1. 創建一個工作簿 (03)
        Workbook workbook = new HSSFWorkbook();
        // 2. 創建一個工作表
        Sheet sheet = workbook.createSheet("狂神觀眾統計表");

        // 3. 創建一個行 (第一行)
        Row row1 = sheet.createRow(0);
        // 4. 創建一個單元格 (1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增觀眾");
        // 創建一個單元格 (1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        // 創建一個行 (第二行)
        Row row2 = sheet.createRow(1);
        // 創建一個單元格 (2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("統計時間");
        // 創建一個單元格 (2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        // 5. 生成一張表 (IO 流)  03 版本, 使用 xls結尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "poi_狂神觀眾統計表03.xls");
        // 6. 輸出
        workbook.write(fileOutputStream);
        // 7. 關閉流
        fileOutputStream.close();

        // 文字提示
        System.out.println("poi_狂神觀眾統計表03.xls 生成完畢 !!");
    }

    @Test
    public void testWrite07() throws Exception {
        // 1. 創建一個工作簿 (07)
        Workbook workbook = new XSSFWorkbook();
        // 2. 創建一個工作表
        Sheet sheet = workbook.createSheet("狂神觀眾統計表");

        // 3. 創建一個行 (第一行)
        Row row1 = sheet.createRow(0);
        // 4. 創建一個單元格 (1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增觀眾");
        // 創建一個單元格 (1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        // 創建一個行 (第二行)
        Row row2 = sheet.createRow(1);
        // 創建一個單元格 (2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("統計時間");
        // 創建一個單元格 (2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        // 5. 生成一張表 (IO 流)  07 版本, 使用 xlsx結尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "poi_狂神觀眾統計表07.xlsx");
        // 6. 輸出
        workbook.write(fileOutputStream);
        // 7. 關閉流
        fileOutputStream.close();

        // 文字提示
        System.out.println("poi_狂神觀眾統計表07.xlsx 生成完畢 !!");
    }

    @Test
    public void testWrite03BigData() throws Exception {
        // 起始時間
        long begin = System.currentTimeMillis();

        // 1. 創建一個工作簿 (03)
        Workbook workbook = new HSSFWorkbook();
        // 2. 創建一個工作表 (使用 預設工作表名稱)
        Sheet sheet = workbook.createSheet();

        // 3. 創建數據 (創建一個行 & 單元格)
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("Big Data 創建完成");

        // 4. 生成一張表 (IO 流)  03 版本, 使用 xls結尾
        FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite_03_BigData.xls");
        // 5. 輸出
        workbook.write(outputStream);
        // 6. 關閉流
        outputStream.close();

        // 結束時間
        long end = System.currentTimeMillis();
        // 計算處理時間
        System.out.println("處理時間: "+(double)(end-begin)/1000+" 秒");

        // 文字提示
        System.out.println("testWrite_03_BigData.xls 生成完畢 !!");
    }

    @Test
    public void testWrite07BigData() throws Exception {
        // 起始時間
        long begin = System.currentTimeMillis();

        // 1. 創建一個工作簿 (07)
        Workbook workbook = new XSSFWorkbook();
        // 2. 創建一個工作表 (使用 預設工作表名稱)
        Sheet sheet = workbook.createSheet();

        // 3. 創建數據 (創建一個行 & 單元格)
        for (int rowNum = 0; rowNum < 65537; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("Big Data 創建完成");

        // 4. 生成一張表 (IO 流)  07 版本, 使用 xlsx結尾
        FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite_07_BigData.xlsx");
        // 5. 輸出
        workbook.write(outputStream);
        // 6. 關閉流
        outputStream.close();

        // 結束時間
        long end = System.currentTimeMillis();
        // 計算處理時間
        System.out.println("處理時間: "+(double)(end-begin)/1000+" 秒");

        // 文字提示
        System.out.println("testWrite_07_BigData.xlsx 生成完畢 !!");
    }

    @Test
    public void testWrite07BigDataSuper() throws Exception {
        // 起始時間
        long begin = System.currentTimeMillis();

        // 1. 創建一個工作簿 (07 升級版的實現類 SXSSFWorkbook)
        Workbook workbook = new SXSSFWorkbook();
        // 2. 創建一個工作表 (使用 預設工作表名稱)
        Sheet sheet = workbook.createSheet();

        // 3. 創建數據 (創建一個行 & 單元格)
        for (int rowNum = 0; rowNum < 200000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("Big Data 創建完成");

        // 4. 生成一張表 (IO 流)  07 版本, 使用 xlsx結尾
        FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite_07_BigDataSuper.xlsx");
        // 5. 輸出
        workbook.write(outputStream);
        // 6. 關閉流
        outputStream.close();

        // 清除臨時文件
        ((SXSSFWorkbook) workbook).dispose();

        // 結束時間
        long end = System.currentTimeMillis();
        // 計算處理時間
        System.out.println("處理時間: "+(double)(end-begin)/1000+" 秒");

        // 文字提示
        System.out.println("testWrite_07_BigDataSuper.xlsx 生成完畢 !!");
    }

}
