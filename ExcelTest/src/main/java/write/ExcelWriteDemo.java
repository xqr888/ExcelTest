package write;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class ExcelWriteDemo {
    public static void main(String[] args) {
        try {
            new ExcelWriteDemo().writeBigData07();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //03版本写入
    public void writeExcel03() throws Exception {
        //1. 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //2. 创建工作表
        Sheet sheet = workbook.createSheet("03版本测试");
        //3. 创建行（第一行）
        Row row1 = sheet.createRow(0);
        //4. 创建单元格（1,1）
        Cell cell11 = row1.createCell(0);
        //5. 写入数据
        cell11.setCellValue("商品ID");

        //6. 写入数据（1,2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("商品名称");

        //第二行（2,1）
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue(1);

        //（2,2）
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue("鼠标");

        //生成一张表-IO流
        FileOutputStream outputStream = new FileOutputStream("./03版本测试.xls");
        workbook.write(outputStream);
        //关闭输出流
        outputStream.close();
        System.out.println("表格生成完毕！");
    }

    //07版本写入
    public void writeExcel07() throws Exception {
        //1. 创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //2. 创建工作表
        Sheet sheet = workbook.createSheet("07版本测试");
        //3. 创建行（第一行）
        Row row1 = sheet.createRow(0);
        //4. 创建单元格（1,1）
        Cell cell11 = row1.createCell(0);
        //5. 写入数据
        cell11.setCellValue("商品ID");

        //6. 写入数据（1,2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("商品名称");

        //第二行（2,1）
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue(1);

        //（2,2）
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue("鼠标");

        //生成一张表-IO流
        FileOutputStream outputStream = new FileOutputStream("./07版本测试.xlsx");
        workbook.write(outputStream);
        //关闭输出流
        outputStream.close();
        System.out.println("表格生成完毕！");
    }

    //03版本批量写入
    public void writeBatchData03() throws Exception {
        // 开始时间
        long start = System.currentTimeMillis();
        // 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        // 创建表
        Sheet sheet = workbook.createSheet("03");
        // 写入数据
        for(int rowNum = 0;rowNum<65536;rowNum++){
            Row row = sheet.createRow(rowNum);
            for(int cellNum = 0;cellNum<20;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum+1);
            }
        }
        // 生成表
        FileOutputStream outputStream = new FileOutputStream("./03BatchData.xls");
        workbook.write(outputStream);
        // 关闭流
        outputStream.close();
        System.out.println("表格生成完毕！");
        // 结束时间
        long end = System.currentTimeMillis();
        System.out.println((end-start)/1000);
    }

    //07版本批量写入
    public void writeBatchData07() throws Exception {
        // 开始时间
        long start = System.currentTimeMillis();
        // 创建工作簿
        Workbook workbook = new XSSFWorkbook();
        // 创建表
        Sheet sheet = workbook.createSheet("07");
        // 写入数据
        for(int rowNum = 0;rowNum<65536;rowNum++){
            Row row = sheet.createRow(rowNum);
            for(int cellNum = 0;cellNum<20;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum+1);
            }
        }
        // 生成表
        FileOutputStream outputStream = new FileOutputStream("./07BatchData.xlsx");
        workbook.write(outputStream);
        // 关闭流
        outputStream.close();
        System.out.println("表格生成完毕！");
        // 结束时间
        long end = System.currentTimeMillis();
        System.out.println((end-start)/1000);
    }

    //07版本批量写入
    public void writeBigData07() throws Exception {
        // 开始时间
        long start = System.currentTimeMillis();
        // 创建工作簿
        Workbook workbook = new SXSSFWorkbook();
        // 创建表
        Sheet sheet = workbook.createSheet("07");
        // 写入数据
        for(int rowNum = 0;rowNum<100000;rowNum++){
            Row row = sheet.createRow(rowNum);
            for(int cellNum = 0;cellNum<20;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum+1);
            }
        }
        // 生成表
        FileOutputStream outputStream = new FileOutputStream("./07BigData.xlsx");
        workbook.write(outputStream);
        // 关闭流
        outputStream.close();
        System.out.println("表格生成完毕！");
        // 结束时间
        long end = System.currentTimeMillis();
        System.out.println((end-start)/1000);
    }

}
