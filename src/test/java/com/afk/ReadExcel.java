package com.afk;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) throws IOException, InterruptedException {
        File scr= new File("C:/Users/lenovo/Documents/E/Testing/selenium-project/maven_project/ExcelReadingApachePoi/readFromExcelTest.xlsx");
        FileInputStream fileInputStream= new FileInputStream(scr);
        XSSFWorkbook wb= new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1= wb.getSheetAt(0);
        int rowCount=sheet1.getLastRowNum();
        System.out.println("Total Numbers of Rows where content is = " + rowCount);
        for (int i = 0; i <= rowCount; i++) {
            String username =sheet1.getRow(i).getCell(0).getStringCellValue();
            System.out.println("Username = " + username);
            String password=sheet1.getRow(i).getCell(1).getStringCellValue();
            System.out.println("Password = " + password);
            Thread.sleep(1000);
        }

    }
}
