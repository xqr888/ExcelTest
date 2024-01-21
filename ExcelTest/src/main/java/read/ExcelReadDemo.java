package read;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import util.ExcelUtil;
import entity.Product;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class ExcelReadDemo {

    public static void main(String[] args) {
        try {
            new ExcelReadDemo().getEntitys();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void readExcel03() throws Exception {
        //1. 通过文件流读取Excel工作簿
        FileInputStream inputStream = new FileInputStream("./03版本测试.xls");
        //2. 获取工作簿
        Workbook workbook = new HSSFWorkbook(inputStream);
        //3. 获取表(通过下标的方式来进行读取 或者 可以采用表名来进行读取)
        Sheet sheet = workbook.getSheetAt(0);
        //4. 获取行（采用下标的方式来进行获取）
        Row row = sheet.getRow(0);
        //5. 获取单元格(采用下标的方式)
        Cell cell = row.getCell(1);
        //6. 读取数据
        String data = cell.getStringCellValue();
        System.out.println(data);
        //7. 关闭流
        inputStream.close();
    }

    public void readExcel07() throws Exception {
        //1. 通过文件流读取Excel工作簿
        FileInputStream inputStream = new FileInputStream("./07版本测试.xlsx");
        //2. 获取工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        //3. 获取表(通过下标的方式来进行读取 或者 可以采用表名来进行读取)
        Sheet sheet = workbook.getSheetAt(0);
        //4. 获取行（采用下标的方式来进行获取）
        Row row = sheet.getRow(0);
        //5. 获取单元格(采用下标的方式)
        Cell cell = row.getCell(1);
        //6. 读取数据
        String data = cell.getStringCellValue();
        System.out.println(data);
        //7. 关闭流
        inputStream.close();
    }

    public void readExcelCellType() throws Exception {
        //1. 通过文件流读取Excel工作簿
        FileInputStream inputStream = new FileInputStream("./商品表.xls");
        //2. 获取工作簿
        Workbook workbook = new HSSFWorkbook(inputStream);
        //3. 获取表(通过下标的方式来进行读取 或者 可以采用表名来进行读取)
        Sheet sheet = workbook.getSheetAt(0);

        //获取标题内容（获取表中的第一行的数据）
        Row title = sheet.getRow(0);
        //非空判断
        if(title!=null){
            //获取标题的单元格数量，用于遍历获取所有单元格
            int cellNum = title.getPhysicalNumberOfCells();
            System.out.println(cellNum);
            for (int i = 0; i < cellNum; i++) {
                //获取所有单元格
                Cell cell = title.getCell(i);
                if(cell!=null){
                    //获取单元格中的数据
                    String value = cell.getStringCellValue();
                    System.out.println(value);
                }
            }
        }

        //获取标题以下的具体内容
        //获取一共有多少行数据
        int rowNum = sheet.getPhysicalNumberOfRows();
        System.out.println(rowNum);
        //跳过第一行标题数据获得以下具体内容
        for (int i = 1; i < rowNum; i++) {
            Row row = sheet.getRow(i);
            if(row!=null){
                //获取每一行里面有多少单元格
                int cellNum = row.getPhysicalNumberOfCells();
                //遍历每一行里面单元格的数据
                for (int j = 0; j < cellNum; j++) {
                    Cell cell = row.getCell(j);
                    if(cell!=null){
                        /**
                         * CellType中定义了不同的枚举类型，来作为表格数据的接收类型
                         * _NONE 未知类型
                         * NUMERIC 数值类型（整数、小数、日期）
                         * STRING 字符串
                         * FORMULA 公式
                         * BLANK 空字符串（没有值），但是有单元格样式
                         * BOOLEAN 布尔值
                         * ERROR 错误单元格
                         */
                        // 获取所有读取数据的类型
                        CellType cellType = cell.getCellType();
                        String cellVal = "";
                        // 根据不同的类型来读取数据
                        switch (cellType){
                            case STRING: //字符串
                                cellVal = cell.getStringCellValue();
                                System.out.println("字符串类型");
                                break;
                            case NUMERIC://数值类型
                                //判断是否为日期类型
                                if(DateUtil.isCellDateFormatted(cell)){
                                    System.out.println("日期类型");
                                    Date date = cell.getDateCellValue();
                                    cellVal = new SimpleDateFormat("yyyy-MM-dd").format(date);
                                }else {
                                    cellVal = cell.toString();
                                    System.out.println("数值类型");
                                }
                                break;
                            case BLANK://空字符串
                                System.out.println("空白字符");
                                break;
                            case BOOLEAN://布尔类型
                                cellVal = String.valueOf(cell.getBooleanCellValue());
                                System.out.println("布尔类型");
                                break;
                            case ERROR://错误类型
                                System.out.println("格式错误");
                                break;
                        }
                        System.out.println(cellVal);
                    }
                }
            }
        }
        inputStream.close();
    }

    public void getFormula() throws Exception{
        //1. 通过文件流读取Excel工作簿
        FileInputStream inputStream = new FileInputStream("./读取公式.xlsx");
        //2. 获取工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        //3. 获取表(通过下标的方式来进行读取 或者 可以采用表名来进行读取)
        Sheet sheet = workbook.getSheetAt(1);

        //获取公式行
        Row row = sheet.getRow(2);
        Cell cell = row.getCell(0);

        System.out.println(cell.getNumericCellValue());

        //获取计算公式
        FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        //获取单元格内容
        CellType cellType = cell.getCellType();
        switch (cellType){
            case FORMULA://公式
                //获取公式
                String formula = cell.getCellFormula();
                System.out.println(formula);
                //获取计算结果
                CellValue value = formulaEvaluator.evaluate(cell);
                String val = value.formatAsString();
                System.out.println(val);
                break;
        }
        inputStream.close();
    }

    public void getEntitys() throws Exception {
        FileInputStream inputStream = new FileInputStream("./商品表.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        List<Product> list = ExcelUtil.readExcel(workbook,inputStream,Product.class);
        System.out.println(list);
    }


}
