package cn.ldp.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

public class poiTest {
    @Test
    void testWriteXMLS() throws Exception {
        Class.forName("com.mysql.cj.jdbc.Driver");
        Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/easyui-test?useUnicode=true&characterEncoding=UTF-8&serverTimezone=GMT%2B8", "root", "root");
        PreparedStatement preparedStatement = connection.prepareStatement("select * from student");
        ResultSet resultSet = preparedStatement.executeQuery();
        //获取有多少个字段
        int columnCount = resultSet.getMetaData().getColumnCount();
//        System.out.println(columnCount);
        //创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建页
        Sheet sheet1 = workbook.createSheet("表单一");
        //初始化excel中的行
        int row = 0;
        while (resultSet.next()){
//            创建excel第row行
            Row row1 = sheet1.createRow(row);
//            初始化excel列
            int j = 0;
            //判断excel当前列和数据库字段数
            while (j < columnCount){
                //创建excel第row行第j+1列
                Cell cell = row1.createCell(j);
//                System.out.println("下面是第" + (row+1) + "行数据第" + (j+1) + "列数据");
//                获取数据库该条记录的第j+1个字段的值
                Object object = resultSet.getObject(j+1);
                //判断字段数据类型并将其复制到excel单元格中
                if (object instanceof Integer){
                    Integer integer = (Integer) object;
                    cell.setCellValue(integer);
//                    System.out.println(integer);
                }else if(object instanceof Date){
                    java.util.Date date = (java.util.Date) object;
                    cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(date));
//                    System.out.println(date);
                }else {
                    cell.setCellValue((String)object);
//                    System.out.println(object);
                }
                j++;
            }
            row ++;
        }
        connection.close();
        Sheet sheet2 = workbook.createSheet("表单二");
        Row row2 = sheet2.createRow(1);
        row2.createCell(0).setCellValue(1);
        row2.createCell(1).setCellValue("哈哈");
        row2.createCell(2).setCellValue(false);
        row2.createCell(3).setCellValue('哈');
        row2.createCell(4).setCellValue(22.2221);
        FileOutputStream outputStream = new FileOutputStream("student.xls");
        workbook.write(outputStream);
        outputStream.close();
    }
    @Test
    void testReadXMLS() throws Exception {
        HSSFWorkbook workbook=new HSSFWorkbook(new FileInputStream(new File("student.xls")));
        //获取excel第一页
        HSSFSheet sheet = workbook.getSheetAt(0);
//        遍历每一行
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            HSSFRow row = sheet.getRow(i);
            //遍历每一列
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                HSSFCell cell = row.getCell(j);
                //获取单元格数据类型
                CellType cellTypeEnum = cell.getCellTypeEnum();
                switch (cellTypeEnum){
                    //数字类型
                    case NUMERIC:
                        //获取double单元格值
                        double numericCellValue = cell.getNumericCellValue();
                        //进行去.0操作
                        DecimalFormat decimalFormat = new DecimalFormat("###################.###########");
                        //转为字符串
                        String formatValue = decimalFormat.format(numericCellValue);
                        System.out.println("第" + (i+1) + "行第" + (j+1) + "列数据是:" + formatValue);
                        break;
                    default:
                        System.out.println(cellTypeEnum);
                        System.out.println("第" + (i+1) + "行第" + (j+1) + "列数据是:" + row.getCell(j));
                }
                }
            }
        }
    }