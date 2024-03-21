package com.dbx.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.dbx.Data;
import com.dbx.DataListener;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class FormulaEvalTest {
    public static void main(String[] args) {
        String baseDir = "C:\\Users\\Administrator\\Desktop\\公式转值";
        String templateDir = "模板";
        String templatePath = baseDir + File.separator + templateDir + File.separator + "template.xlsx";
        String suffix = ".xlsx";
        String outputPath = "C:\\Users\\Administrator\\Desktop\\test.xlsx";
        String dependExcelPath = "C:\\JL.xlsx";
//        Map<String, Object> map = new HashMap<>();
//        map.put("code", "1312");
//        EasyExcel.write(outputPath).withTemplate(templatePath).sheet().doFill(map);

//        try (ExcelWriter build = EasyExcel.write(outputPath).withTemplate(templatePath).build();) {
//            WriteSheet writeSheet = EasyExcel.writerSheet().build();
//            Map<String, Object> map = new HashMap<>();
//            map.put("code", "1312");
//            build.fill(map, writeSheet);
//        }

        // 这里使用 apache poi 读取文件进行公式转值
        try (FileInputStream fis = new FileInputStream(outputPath);
             FileInputStream dependFis = new FileInputStream(dependExcelPath);
             FileOutputStream fos = new FileOutputStream(outputPath + "_eval")) {
            File outputFile = new File(outputPath);
            File dependFile = new File(dependExcelPath);

            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFWorkbook dependWb = new XSSFWorkbook(dependFis);

            XSSFSheet sheet = workbook.getSheetAt(0);
            sheet.setForceFormulaRecalculation(true);
            XSSFFormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator(workbook);
            XSSFFormulaEvaluator dependFe = new XSSFFormulaEvaluator(dependWb);

            // 设置 workbook 的运行环境
            String[] workbookNames = { toFileSchemePath(outputFile), toFileSchemePath(dependFile)};
            XSSFFormulaEvaluator[] evaluators = {formulaEvaluator, dependFe};
            XSSFFormulaEvaluator.setupEnvironment(workbookNames, evaluators);


            Iterator<Row> rowIter = sheet.iterator();
            while(rowIter.hasNext()) {
                Row row = rowIter.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if(CellType.FORMULA.equals(cell.getCellType())) {
                        // 尝试对公式进行计算 并 赋值到 单元格中
                        formulaEvaluator.evaluateInCell(cell);
                    }
                }
            }

            workbook.write(fos);
        } catch (Exception e) { e.printStackTrace(); }
    }

    public static String toFileSchemePath(File file) {
        // 判断是否是C盘下，excel下 C盘下的路径是直接使用根目录"/"作为替代
        String path = file.toURI().getPath();
        if(path.startsWith("/C:")) {
            return path.substring("/C:".length());
        }

        // 如果不是C盘下的路径，直接拼接 file:// 协议头返回
        return "file://" + path;
    }
}
