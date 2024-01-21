package util;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 解析Excel表格工具类
 */
public class ExcelUtil {
    /**
     * 目的：用户只需要传入Workbook对象（匹配版本），文件输入流，对应实体类Class，
     * 就可以得到解析表格以后的结果，同时通过传入的实体类类型集合的方式来返回
     */
    public static <T> List<T> readExcel(Workbook workbook, FileInputStream inputStream,Class<T> clazz) throws Exception{
        //给用户返回的实体类集合
        List<T> result = new ArrayList<>();
        // 在工作簿中获取目标工作表
        Sheet sheet = workbook.getSheetAt(0);
        // 获取工作表中的行数
        int rowNum = sheet.getPhysicalNumberOfRows();
        // 获取第一行数据（隐藏行）
        Row row = sheet.getRow(1);
        // 遍历第一行数据，宾利出的数据就是当前实体类对应的所有属性，同时要把这些数据放入到Map中的key
        List<String> key = new ArrayList<>();
        // 具体遍历
        for(Cell cell : row){
            if(cell!=null){
                //获取单元格中的数据
                String value = getCellValue(cell);
                key.add(value);
                System.out.println(value);
            }
        }

        // 遍历正式数据
        for (int i = 2; i < rowNum; i++) {
            //获取属性名以下的数据
            row = sheet.getRow(i);
            if(row != null){
                // 计数器 j 用于映射数据使用
                int j = 0;
                // 用于保存每条数据的Map，并且在Map中建立属性对应数据的映射关系
                Map<String,String> excelMap = new HashMap<>();
                for(Cell cell : row){
                    if(cell != null){
                        // 把所有单元格中的数据格式设置为String
                        String value = getCellValue(cell);
                        if(value != null && !value.equals("")){
                            //将每个单元格的数据存储到集合中
                            excelMap.put(key.get(j),value);
                            j++;
                        }
                    }
                }
                // 创建对应实体类类型，并且把读取到的数据转换为实体类对象
                T t = mapToEntity(excelMap,clazz);
                result.add(t);
            }
        }
        inputStream.close();
        return result;
    }

    public static NumberFormat nf = NumberFormat.getNumberInstance();
    static {
        nf.setGroupingUsed(false);
    }

    public static String getCellValue(Cell cell){
        String cellVal = "";
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
            // 根据不同的类型来读取数据
            switch (cellType){
                case STRING: //字符串
                    cellVal = cell.getStringCellValue();
                    break;
                case NUMERIC://数值类型
                    //判断是否为日期类型
                    if(DateUtil.isCellDateFormatted(cell)){
                        Date date = cell.getDateCellValue();
                        cellVal = new SimpleDateFormat("yyyy-MM-dd").format(date);
                    }else {
                        cellVal = nf.format(cell.getNumericCellValue());
                    }
                    break;
                case BLANK://空字符串
                    break;
                case BOOLEAN://布尔类型
                    cellVal = String.valueOf(cell.getBooleanCellValue());
                    break;
                case ERROR://错误类型
                    break;
            }
            System.out.println(cellVal);
        }
        return cellVal;
    }

    private static <T> T mapToEntity(Map<String,String> map,Class<T> entity){
        T t = null;
        try {
            t = entity.newInstance();
            for(Field field : entity.getDeclaredFields()){
                if(map.containsKey(field.getName())){
                    boolean flag = field.isAccessible();
                    field.setAccessible(true);
                    //获取Map中的属性对应的值
                    String str = map.get(field.getName());
                    // 获取实体类属性的类型
                    String type = field.getGenericType().toString();
                    // 重新制定对应属性的值
                    if(str != null) {
                        if (type.equals("class java.lang.String")) {
                            field.set(t, str);
                        } else if (type.equals("class java.lang.Double")) {
                            field.set(t, Double.parseDouble(String.valueOf(str)));
                        } else if (type.equals("class java.lang.Integer")) {
                            field.set(t, Integer.parseInt(String.valueOf(str)));
                        } else if (type.equals("class java.util.Date")) {
                            Date date = new SimpleDateFormat("yyyy-MM-dd").parse(str);
                            field.set(t, date);
                        }
                    }
                    field.setAccessible(flag);
                }
            }
            return t;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return t;
    }
}

