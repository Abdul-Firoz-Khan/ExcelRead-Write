package com.afk;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class WriteExcel {
    public static void main(String[] args) throws IOException {
        File src = new File("C:/Users/lenovo/Documents/E/Testing/selenium-project/maven_project/ExcelReadingApachePoi/readFromExcelTest.xlsx");
        FileInputStream fileInputStream= new FileInputStream(src);
        XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = wb.getSheetAt(0);

        sheet1.getRow(0).createCell(2).setCellValue("Fail");
        sheet1.getRow(1).createCell(2).setCellValue("Pass");
        FileOutputStream fileOutputStream= new FileOutputStream(src);
        wb.write(fileOutputStream);


    }
}
