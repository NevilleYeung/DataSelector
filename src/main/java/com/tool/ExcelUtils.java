package com.tool;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class ExcelUtils {
    /**
     * getWorkBook
     *
     * @param titleList
     * @param contentlist
     * @param sheetName
     * @return xwk
     */
    public XSSFWorkbook getWorkBook(List<String> titleList, List<List<String>> contentlist, String sheetName) {
        XSSFWorkbook xwk = new XSSFWorkbook();
        XSSFDataFormat format = xwk.createDataFormat();
        XSSFCellStyle cellStyle = xwk.createCellStyle();
        XSSFSheet xssfSheet = xwk.createSheet(sheetName);
        cellStyle.setDataFormat(format.getFormat("@"));//文本格式
        int j = 0;
        createHeader(xssfSheet, cellStyle, titleList, j);
        int size = contentlist.size();
        for (j = 0; j < size; j++) {
            List<String> oneRow = contentlist.get(j);
            createContent(xssfSheet, cellStyle, oneRow, j);
            oneRow = null;
        }
        return xwk;
    }

    /**
     * createHeader
     *
     * @param xssfSheet
     * @param titleList
     */
    private void createHeader(XSSFSheet xssfSheet, XSSFCellStyle cellStyle, List<String> titleList, int j) {
        XSSFRow rowTitle = xssfSheet.createRow(j);
        for (int cellTitle = 0; cellTitle < titleList.size(); cellTitle++) {
            Cell cellIndex = rowTitle.createCell(cellTitle);
            cellIndex.setCellStyle(cellStyle);
            cellIndex.setCellValue(titleList.get(cellTitle));
        }
    }

    /**
     * createHeader
     *
     * @param xssfSheet
     * @param oneRow
     * @param j
     */
    private void createContent(XSSFSheet xssfSheet, XSSFCellStyle cellStyle, List<String> oneRow, int j) {
        XSSFRow rowContent = xssfSheet.createRow(j + 1);
        for (int cellContent = 0; cellContent < oneRow.size(); cellContent++) {
            Cell cellIndex = rowContent.createCell(cellContent);
            cellIndex.setCellStyle(cellStyle);
            cellIndex.setCellValue(oneRow.get(cellContent));
        }
    }
}
