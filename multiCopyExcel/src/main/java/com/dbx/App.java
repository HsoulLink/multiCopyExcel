package com.dbx;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

/**
 * Hello world!
 *
 */
public class App 
{
    private static String suffix = ".xlsx";
    private static String baseDir = "C:\\Users\\Administrator\\Desktop\\刨槽";
    private static String templateDir = "模板";
    private static String templateFileName  = "template.xlsx";
    private static String templatePath =
                                    new StringBuilder(baseDir)
                                            .append(File.separator)
                                            .append(templateDir)
                                            .append(File.separator)
                                            .append(templateFileName).toString();
    private static String totalExcelDir = "总表";
    private static String totalExcelName = "LW.xlsx";
    private static String totalExcelFilePath =
                                new StringBuilder(baseDir)
                                        .append(File.separator)
                                        .append(totalExcelDir)
                                        .append(File.separator)
                                        .append(totalExcelName).toString();

    private static String outputDir = "LW";
    private static String outputPath =
                                new StringBuilder(baseDir)
                                        .append(File.separator)
                                        .append(outputDir).toString();

    private static boolean isEvalFormula = true;

    public static void main( String[] args ) throws IOException {
        // 读取刨槽总表的信息
        DataListener dataListener = new DataListener();
        EasyExcel.read(totalExcelFilePath, Data.class, dataListener).sheet().doRead();

        // 批量写入
        String destPath = null;
        for (Data data : dataListener.getDataList()) {
            if(StringUtils.isBlank(data.getFirstPit())) {
                // 说明具体刨槽内容还不能确定，跳过
                continue;
            }
            // 1.填充对应的excel数据
            destPath = new StringBuilder(outputPath).append(File.separator).append(data.getCode()).append(suffix).toString();
            EasyExcel.write(destPath).withTemplate(templatePath).sheet().doFill(data);

            // 2.计算公式并把结果值填充到单元格上
            if(isEvalFormula) {
                evalFormula(destPath);
            }
        }
    }

    public static void evalFormula(String destPath) {
        // 这里使用 apache poi 读取文件进行公式转值
        // 1. 读取目标文件
        XSSFWorkbook workbook = null;
        XSSFWorkbook depWorkbook = null;
        try (FileInputStream fis = new FileInputStream(destPath);
             FileInputStream dependFis = new FileInputStream(totalExcelFilePath);
        ){
            workbook = new XSSFWorkbook(fis);
            depWorkbook = new XSSFWorkbook(dependFis);
        } catch (Exception e) { e.printStackTrace(); }

        // 获取需要进行计算的excel的第一个 sheet

        XSSFSheet sheet = workbook.getSheetAt(0);
        sheet.setForceFormulaRecalculation(true);
        // 2. 设置 workbook 的运行环境。(由于需要计算的表有依赖总表，所以需要配置其依赖)
        // apache poi 使用FormulaEvaluator进行计算
        XSSFFormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator(workbook);
        XSSFFormulaEvaluator depFormulaEvaluator = new XSSFFormulaEvaluator(depWorkbook);
        String[] workbookNames = { toFileSchemePath(new File(destPath)), toFileSchemePath(new File(totalExcelFilePath))};
        XSSFFormulaEvaluator[] evaluators = {formulaEvaluator, depFormulaEvaluator};
        XSSFFormulaEvaluator.setupEnvironment(workbookNames, evaluators);

        // 3. 计算单元格类型为 FORMULA(公式) 的值并填充到原单元格中
        Iterator<Row> rowIter = sheet.iterator();
        while(rowIter.hasNext()) {
            Row row = rowIter.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(CellType.FORMULA.equals(cell.getCellType())) {
                    //TODO: 由于总表是根据 厂家 名称分开的，因此还需要修改一下公式链接的文件路径。目前暂时找不到怎么修改它依赖的路径

                    // 尝试对公式进行计算 并 赋值到 单元格中
                    formulaEvaluator.evaluateInCell(cell);
                }
            }
        }

        // 4. 将处理完的结果写回到文件
        try(FileOutputStream outFos = new FileOutputStream(destPath)) {
            workbook.write(outFos);
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
