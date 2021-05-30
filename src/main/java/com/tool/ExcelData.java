package com.tool;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class ExcelData {
    private XSSFSheet sheet;

    /**
     * 构造函数，初始化excel数据
     *
     * @param filePath  excel路径
     * @param sheetName sheet表名
     */
    ExcelData(String filePath, String sheetName) {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(filePath);
            XSSFWorkbook sheets = new XSSFWorkbook(fileInputStream);
            //获取sheet
            sheet = sheets.getSheet(sheetName);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 根据行和列的索引获取单元格的数据
     *
     * @param row
     * @param column
     * @return
     */
    public String getExcelDateByIndex(int row, int column) {
        XSSFRow row1 = sheet.getRow(row);
        String cell = row1.getCell(column).toString();
        return cell;
    }

    /**
     * 根据某一列值为“******”的这一行，来获取该行第x列的值
     *
     * @param caseName
     * @param currentColumn 当前单元格列的索引
     * @param targetColumn  目标单元格列的索引
     * @return
     */
    public String getCellByCaseName(String caseName, int currentColumn, int targetColumn) {
        String operateSteps = "";
        // 获取行数
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rows; i++) {
            XSSFRow row = sheet.getRow(i);
            String cell = row.getCell(currentColumn).toString();
            if (cell.equals(caseName)) {
                operateSteps = row.getCell(targetColumn).toString();
                break;
            }
        }
        return operateSteps;
    }

    // 打印excel数据
    public void readExcelData() {
        //获取行数
        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rows; i++) {
            //获取列数
            XSSFRow row = sheet.getRow(i);
            int columns = row.getPhysicalNumberOfCells();
            for (int j = 0; j < columns; j++) {
                String cell = row.getCell(j).toString();
                System.out.println(cell);
            }
        }
    }

    /**
     * 获取行数
     *
     * @return 行数
     */
    public int getPhysicalNumberOfRows() {
        return sheet.getPhysicalNumberOfRows();
    }

    /**
     * 获取一行的数据
     *
     * @param index Index
     * @return 一行的数据
     */
    public XSSFRow getRow(int index) {
        return sheet.getRow(index);
    }

    /**
     * 只获取两列的数据
     *
     * @return 文件内前两列的数据
     */
    public double[][] getData() {
        /*
        数据样例：
        位移   压力
        0.013 -294.523
        0.014 -288.367
        0.015 -282.464
        0.016 -276.198
        0.017 -270.121
         */
        // 获取行数
        int rows = sheet.getPhysicalNumberOfRows();
        double[][] datas = new double[rows - 1][2];

        // 第一列是属性名，跳过
        for (int i = 1; i < rows; i++) {
            // 获取列数
            XSSFRow row = sheet.getRow(i);
            if (row == null) continue;
            for (int j = 0; j < 2; j++) {
                String cell = row.getCell(j).toString();
                datas[i - 1][j] = Double.parseDouble(cell);
            }
        }

        return datas;
    }
}
