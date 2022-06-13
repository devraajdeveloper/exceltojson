package com.raaj.excelToCSV.controller;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
public class ExcelController {

    @PostMapping("/upload")
    public String excelReader(@RequestParam("file")MultipartFile excel, @RequestParam("country_code") String countryCode, @RequestParam("sheet_name") String sheetName){
        String key = null;
        String value = null;
        StringBuilder main = new StringBuilder();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(excel.getInputStream());
            XSSFSheet sheet = workbook.getSheetAt(workbook.getSheetIndex(sheetName));

            for(int i=1; i<sheet.getPhysicalNumberOfRows();i++) {
                XSSFRow row = sheet.getRow(i);
                for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
                    if(j==0){
                        key = "\""+row.getCell(j)+"\"" +":";
                        System.out.print("\""+row.getCell(j)+"\"" +":");
                    }
                    else if(j==2) {
                        value = "\""+row.getCell(j)+"\"";
                        System.out.print("\""+row.getCell(j)+"\"");
                    }
                }
                System.out.println("");
                main.append("   ").append(key).append(value).append("\n");
            }
            System.out.println(main);

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        String finalData = "var "+countryCode+" = "+"{"+"\n"+main+"   }";

        return finalData;
    }
}
