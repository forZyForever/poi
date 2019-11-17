package com.itheima;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class POIDemo {

    public static void main(String[] args) throws IOException {
        //readExcel();
        writeExcel();
    }

    private static void writeExcel() throws IOException {
        // 创建工作簿
        try (
                XSSFWorkbook wk = new XSSFWorkbook();
        ) {
            // 创建工作表
            XSSFSheet sht = wk.createSheet("测试工作表");
            // 创建行
            XSSFRow row = sht.createRow(0);
            // 创建单元格
            XSSFCell cell = row.createCell(0);
            //  给单元格赋值
            cell.setCellValue("姓名");
            row.createCell(1).setCellValue("年龄");
            row.createCell(2).setCellValue("性别");

            // 表格内容,
            row = sht.createRow(1);
            row.createCell(0).setCellValue("张三");
            row.createCell(1).setCellValue(23);
            row.createCell(2).setCellValue("男");
            //....
            // 保存下来
            wk.write(new FileOutputStream(new File("d:\\userInfo2.xlsx")));

        }
    }

    private static void readExcel() throws IOException {
        // 创建工作簿 try(实现closeable接口), finally{close()}
        try (
                XSSFWorkbook wk = new XSSFWorkbook("d:\\userInfo.xlsx");
        ) {
            // 获取工作表
            XSSFSheet sht = wk.getSheetAt(0);
            // 获取行 下标从0开始
            XSSFRow row = null;
            int lastRowNum = sht.getLastRowNum();// 获取工作表中最后一行的行号, 遍历时用这个
            //sht.getPhysicalNumberOfRows();// 一共有多少行
            for (int i = 0; i < lastRowNum; i++) {
                row = sht.getRow(i);
                // 获取单元格 下标从0开始
                // row.getCell(0);
                // row.getPhysicalNumberOfCells();// 一共多少个单元格
                // row.getLastCellNum();// 最后一个单元格的编号
                for (Cell cell : row) {
                    // 判断单元格的类型
                    int cellType = cell.getCellType();
                    // 获取单元格内容
                    if (cellType == Cell.CELL_TYPE_NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + ",");
                    } else {
                        System.out.print(cell.getStringCellValue() + ",");
                    }
                }
                System.out.println();
            }
        }
    }
}
