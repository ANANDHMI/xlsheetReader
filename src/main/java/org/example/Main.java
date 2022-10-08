package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws IOException {

        final writeValues createNewSheet = new writeValues();
        final Map<String, Double> mapping = new HashMap<>();

        final List<String> subjectList = new ArrayList<>();


        final List<Double> markList = new ArrayList<>();


        File file = new File("source of truth file path");
        try {
            FileInputStream inputStream = new FileInputStream(file);

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sh = workbook.getSheet("Sheet1");

            for (int r = 1; r <= sh.getLastRowNum(); r++) {

                String subjectKey = sh.getRow(r)
                        .getCell(0)
                        .getStringCellValue();

                Double markValue = (double) sh.getRow(r)
                        .getCell(1)
                        .getNumericCellValue();

                subjectList.add(subjectKey);
                markList.add(markValue);

                mapping.put(subjectKey, 0.0);

            }


            List<String> valuesListCopy = new ArrayList(mapping.keySet());

            for (int i = 0; i < valuesListCopy.size(); i++) {
                for (int j = 0; j < subjectList.size(); j++) {
                    if (valuesListCopy.get(i) == subjectList.get(j)) {

                        mapping.replace(valuesListCopy.get(i), mapping.get(valuesListCopy.get(i)) + markList.get(j));
                    }
                }
            }
            workbook.close();
            inputStream.close();
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.print("keys : ");
        System.out.println(subjectList);
        System.out.print("values : ");
        System.out.println(markList);
        System.out.print("Extracted Key values without duplicates : ");
        System.out.println(mapping);
        //This method creates new sheet and insert the map used in
        createNewSheet.createSheet(mapping);


    }

    //read values with respect to data type inside Excel sheet
    public static void printCellValue(Cell cell) {
        CellType cellType = cell.getCellType().equals(CellType.FORMULA)
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (cellType.equals(CellType.STRING)) {
            System.out.print(cell.getStringCellValue() + " | ");
        }
        if (cellType.equals(CellType.NUMERIC)) {
            if (DateUtil.isCellDateFormatted(cell)) {
                System.out.print(cell.getDateCellValue() + " | ");
            } else {
                System.out.print(cell.getNumericCellValue() + " | ");
            }
        }
        if (cellType.equals(CellType.BOOLEAN)) {
            System.out.print(cell.getBooleanCellValue() + " | ");
        }
    }

}
