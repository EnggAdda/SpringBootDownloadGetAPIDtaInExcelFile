package com.example.springbootdownloadexcelfile.util;

import com.example.springbootdownloadexcelfile.entity.Product;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelUtil {

    public static String Header[] = {"id", "name", "price", "quantity",};

    public static String SHEET_NAME = "Sheet";
    public static String SHEET_NAME1 = "Sheet1";

    public static ByteArrayInputStream dataToExcel(List<Product> productList) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
      try{

          //create workbook


          Sheet sheet = workbook.createSheet(SHEET_NAME);
          Sheet sheet1 = workbook.createSheet(SHEET_NAME1);
          Row row  =sheet.createRow(0);

          for(int  i =0; i< Header.length; i++){
              Cell cell = row.createCell(i);
              cell.setCellValue(Header[i]);


          }
          Row row1  =sheet1.createRow(0);

          for(int  i =0; i< Header.length; i++){
              Cell cell = row1.createCell(i);
              cell.setCellValue(Header[i]);


          }

          //value row

          int rowIndex  =1;

          for(Product  p  : productList){
              Row dataRow =sheet.createRow(rowIndex);
              Row dataRow1 =sheet1.createRow(rowIndex);

              rowIndex ++;
              dataRow.createCell(0).setCellValue(p.getId());
              dataRow.createCell(1).setCellValue(p.getName());
              dataRow.createCell(2).setCellValue(p.getPrice());
              dataRow.createCell(3).setCellValue(p.getQuantity());

              dataRow1.createCell(0).setCellValue(p.getId());
              dataRow1.createCell(1).setCellValue(p.getName());
              dataRow1.createCell(2).setCellValue(p.getPrice());
              dataRow1.createCell(3).setCellValue(p.getQuantity());



          }

          workbook.write(out);
          return new ByteArrayInputStream(out.toByteArray());
      }


       catch (Exception e){
           return null;

    }
      finally {
          workbook.close();
          out.close();
      }


    }
}
