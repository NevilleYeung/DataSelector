package com.tool;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 选取数据的工具
 */
public class Selector {
    // 位移增量，单位：mm（毫米）
    private static final BigDecimal DISPLACEMENT_INCREMENT = new BigDecimal( 0.01);
    // 压力增量，单位：N
    private static final BigDecimal PRESSURE_INCREMENT = new BigDecimal( 6);

    private static final String SHEET_NAME = "Sheet1";

    private static final String INPUT_DIR = "D:\\Input\\";
    private static final String OUTPUT_DIR = "D:\\Output\\";
    private static final String EXCEL_SUFFIX = ".xlsx";
    private static final String TEMP_FILE_HEAD = "~";

    public static void main(String[] args) throws Exception {
        // 检查、创建目录
        createDir();
        System.out.println("select data begin...");

        // 获取输入目录下的所有excel文件
        String[] inpuFileNames = new File(INPUT_DIR).list(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                // 只获取.xlsx后缀名的文件，排除临时文件
                if (name.endsWith(EXCEL_SUFFIX) && !name.startsWith(TEMP_FILE_HEAD)) {
                    return true;
                }
                return false;
            }
        });

        if (inpuFileNames == null || inpuFileNames.length == 0) {
            System.out.println("no input files in D:\\Input\\ ...");
            return;
        }

        for (String inputFile: inpuFileNames) {
            try {
                ExcelData sheet1 = new ExcelData(INPUT_DIR + inputFile, SHEET_NAME);

                // 从文件读取数据
                double[][] inputDatas = sheet1.getData();
                if (inputDatas == null || inputDatas.length == 0) {
                    System.out.println("inputFile: " + inputFile + " 数据是空");
                    return;
                }

                // 按要求过滤数据
                List<List<String>> contentsList = filterExcelData(inputDatas);

                // 写入文件的内容
                List<String> titleList = new ArrayList<String>();
                titleList.add("位移");
                titleList.add("压力");
//                List<List<String>> contentsList = new ArrayList<List<String>>();
//                for (Data data : contentsList) {
//                    contentsList.add(Arrays.asList(String.valueOf(data.displacement), String.valueOf(data.pressure)));
//                }

                // 将挑选出来的数据，写入输出文件
                String fileName = OUTPUT_DIR + inputFile;
                writeData2excel(titleList, contentsList, fileName);
            } catch (Throwable t) {
                System.out.println("handle " + inputFile + " failed, " + t);
            }
        }

        System.out.println("select data done...");
    }

    private static void createDir() throws Exception {
        File file = new File(OUTPUT_DIR);
        if (!file.exists()) {
            file.mkdirs();
            return;
        }

        if (!file.isDirectory()) {
            throw new Exception(OUTPUT_DIR + "不是个文件夹");
        }

        if (file.list().length > 0) {
            throw new Exception(OUTPUT_DIR + "不是个空文件夹");
        }
    }

    /**
     * 按要求过滤数据
     *
     * @param inputDatas inputDatas
     * @return 返回
     */
    private static List<List<String>> filterExcelData(double[][] inputDatas) {
        // 数据选取要求：位移增量达到0.01mm 或 压力增量达到6N
        // 注意：与上一次符合条件的数据进行比较
        int left = 0;
        int right = left + 1;
        BigDecimal lastDis = BigDecimal.ZERO;
        BigDecimal lastPres = BigDecimal.ZERO;

        List<List<String>> contentsList = new ArrayList<List<String>>();

        for (int i = 0; i < inputDatas.length; i++) {
            BigDecimal tmpDis = BigDecimal.valueOf(inputDatas[i][0]);
            BigDecimal tmpPres = BigDecimal.valueOf(inputDatas[i][1]);
            int disInc = tmpDis.subtract(lastDis).abs().compareTo(DISPLACEMENT_INCREMENT);
            int presInc = tmpPres.subtract(lastPres).abs().compareTo(PRESSURE_INCREMENT);
            if (disInc >= 0 || presInc >= 0) {
                contentsList.add(Arrays.asList(String.valueOf(inputDatas[i][0]), String.valueOf(inputDatas[i][1])));
                lastDis = tmpDis;
                lastPres = tmpPres;
            }
        }

        return contentsList;
    }

    private static void writeData2excel(List<String> titleList, List<List<String>> contentsList, String fileName) {
        XSSFWorkbook workBook = null;
        FileOutputStream output = null;
        try {
            ExcelUtils eu = new ExcelUtils();
            workBook = eu.getWorkBook(titleList, contentsList, SHEET_NAME);
            output = new FileOutputStream(fileName);
            workBook.write(output);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (output != null) {
                    output.flush();
                    output.close();
                }
                if (workBook != null) {
                    workBook.close();
                }
            }
            catch (Throwable t) {
                t.printStackTrace();
            }
        }
    }
}
