package org.example;// Java Program to Illustrate Reading
// Data to Excel File Using Apache POI

// Import statements
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Main class
public class Main {
    //-------------------------------citire fisier
    // Main driver method
    public static void main(String[] args) throws IOException {

        // Try block to check for exceptions

        FileInputStream file = new FileInputStream(new File("citire-excel.xlsx"));

//Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(file);

//Get first/desired sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);

//Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();

        List<String> list_coll = new ArrayList<>();
        List<String> list_row = new ArrayList<>();
        Row newrow = sheet.createRow(1);
        Cell newcell = newrow.createCell(4);

        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            //For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                //Check the cell type and format accordingly
                switch (cell.getCellType()) {

                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + " ");
                        list_row.add(String.valueOf(cell.getNumericCellValue()));
                        newcell.setCellValue(cell.getNumericCellValue());
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue() + " ");
                        list_row.add((cell.getStringCellValue()));
                        newcell.setCellValue(cell.getStringCellValue());
                        break;
                    case FORMULA:
                        list_row.add(String.valueOf(cell.getNumericCellValue()));
                        newcell.setCellValue(cell.getNumericCellValue());
                        break;
                }
                for(int i = 0; i < list_row.size(); i++)
                {
                    list_coll.add(list_row.get(i));
                }
                list_row.clear();
            }
            System.out.println(" ");
        }
        file.close();

    }
    //-----------------scriere in fisier


}
