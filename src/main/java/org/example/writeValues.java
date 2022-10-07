package org.example;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class writeValues {

    public void createSheet(final Map<String, Double> mapping) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sheet1");

        // Creating HashMap and entering data

        int rowNo = 0;

        for (HashMap.Entry entry : mapping.entrySet()) {
            XSSFRow row = sheet.createRow(rowNo++);
            row.createCell(0).setCellValue((String) entry.getKey());
            row.createCell(1).setCellValue((Double) entry.getValue());
        }

        FileOutputStream file = new FileOutputStream("new file will be created -->filepath/filename<--");
        workbook.write(file);
        file.close();
        System.out.println("Data Copied to Excel");
    }
}
