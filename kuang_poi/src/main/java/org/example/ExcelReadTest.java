package org.example;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

public class ExcelReadTest {

    String PATH = "..\\..\\";

    @Test
    public void testRead03() throws Exception {

        // 1. 獲取 文件流
        FileInputStream inputStream = new FileInputStream(PATH + "poi_狂神觀眾統計表03.xls");

        // 2. 創建一個工作簿 (03) (文件流)
        Workbook workbook = new HSSFWorkbook(inputStream);
        // 3. 指定一個工作表
        Sheet sheet = workbook.getSheetAt(0);

        // 4. 指定一個行 (第一行)
        Row row1 = sheet.getRow(0);
        // 5. 指定一個單元格 (1,1)
        Cell cell11 = row1.getCell(0);
        System.out.println(cell11.getStringCellValue());

        // 指定一個行 (第二行)
        Row row2 = sheet.getRow(1);
        // 指定一個單元格 (2,1)
        Cell cell21 = row2.getCell(0);
        System.out.println(cell21.getStringCellValue());

        // 6. 關閉流
        inputStream.close();

        // 文字提示
        System.out.println("讀取 poi_狂神觀眾統計表03.xls 完成 !!");
    }

    @Test
    public void testRead07() throws Exception {

        // 1. 獲取 文件流
        FileInputStream inputStream = new FileInputStream(PATH + "poi_狂神觀眾統計表07.xlsx");

        // 2. 創建一個工作簿 (07) (文件流)
        Workbook workbook = new XSSFWorkbook(inputStream);
        // 3. 指定一個工作表
        Sheet sheet = workbook.getSheetAt(0);

        // 4. 指定一個行 (第一行)
        Row row1 = sheet.getRow(0);
        // 5. 指定一個單元格 (1,1)
        Cell cell11 = row1.getCell(0);
        System.out.println(cell11.getStringCellValue());

        // 指定一個行 (第二行)
        Row row2 = sheet.getRow(1);
        // 指定一個單元格 (2,1)
        Cell cell21 = row2.getCell(0);
        System.out.println(cell21.getStringCellValue());

        // 6. 關閉流
        inputStream.close();

        // 文字提示
        System.out.println("讀取 poi_狂神觀眾統計表07.xlsx 完成 !!");
    }

    @Test
    public void testReadDiffCellType03() throws Exception {

        // 1. 獲取 文件流
//        FileInputStream inputStream = new FileInputStream(PATH + "明細表03.xls");
//        FileInputStream inputStream = new FileInputStream(PATH + "2021-0625-ETF1-收盤-trim.xls");
//        FileInputStream inputStream = new FileInputStream(PATH + "2021-0625-ETF1-收盤.xls");

        FileInputStream inputStream = new FileInputStream(PATH + "1027\\20211027.xls");
//        FileInputStream inputStream = new FileInputStream(PATH + "1027\\StockList.xls");
//        FileInputStream inputStream = new FileInputStream(PATH + "1027\\StockList (1).xls");
//        FileInputStream inputStream = new FileInputStream(PATH + "1027\\StockList (2).xls");
//        FileInputStream inputStream = new FileInputStream(PATH + "1027\\StockList (3).xls");

        // 2. 創建一個工作簿 (03) (文件流)
        Workbook workbook = new HSSFWorkbook(inputStream);
        // 3. 指定一個工作表
        Sheet sheet = workbook.getSheetAt(0);

        // 4. 指定一個行 (第一行) (獲取 標題內容)
        Row rowTitle = sheet.getRow(0);
        if (rowTitle!=null) {
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum =0; cellNum < cellCount; cellNum++) {
                HSSFCell cell = (HSSFCell) rowTitle.getCell(cellNum);
                if (cell!=null) {
                    CellType cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
            System.out.println();
        }

        // 5. 獲取 表中的內容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData!=null) {
                // 讀取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("[" + (rowNum+1) + "-" + (cellNum+1) + "]");
                    HSSFCell cell = (HSSFCell) rowData.getCell(cellNum);
                    // 匹配 列的數據類型
                    if (cell!=null) {
                        CellType cellType = cell.getCellType();
                        String cellValue = "";

                        switch ( cellType ) {
                            case STRING:  // 字串
                                System.out.print(" [字串] ");
                                cellValue = cell.getStringCellValue();
                                break;
                            case BOOLEAN:  // 布林
                                System.out.print(" [布林] ");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case BLANK:  // 空
                                System.out.print(" [空 ] ");
                                break;
                            case NUMERIC:  // 數字 (日期 or 普通數字)
//                                System.out.print(" [NUMERIC] ");
                                if (DateUtil.isCellDateFormatted(cell)) {  // 日期
                                    System.out.print(" [日期] ");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                } else {
                                    // 非日期格式, 防止數字過長 (轉換為 字符串輸出)
                                    System.out.print(" [數字] ");
                                    cell.setCellType(CellType.STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case ERROR:
                                // 數據類型 錯誤
                                System.out.print(" [錯誤] ");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
            System.out.println();
        }

        // 6. 關閉流
        inputStream.close();

        // 文字提示
        System.out.println("讀取 明細表03.xls 完成 !!");
    }

}
