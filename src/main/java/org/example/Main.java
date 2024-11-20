package org.example;

import java.io.PrintStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import static java.util.Calendar.DATE;

public class Main {
    public static void main(String[] args) throws IOException {
        /*
        System.out.println((Charset.defaultCharset()));
        String ru = "Русский язык";
        PrintStream ps = new PrintStream(System.out, true, "UTF-8");
        System.out.println(ru.length());
        System.out.println(ru);
        ps.println(ru);*/

        FileInputStream file1 = new FileInputStream(new File("/testExcel/924.xlsx"));
        FileInputStream file2 = new FileInputStream(new File("/testExcel/teacher.xlsx"));


        //Create Workbook instance holding reference to .xlsx file1
        XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
        XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

        //Get first/desired sheet1 from the workbook1
        XSSFSheet sheet1 = workbook1.getSheetAt(0);
        XSSFSheet sheet2 = workbook2.getSheetAt(0);


        int startRow = 0;
        int endRow = 2;
        int startCol = 3;
        int endCol = 6;
        int[] rowArray={11,14,17,20,24,27,30,33,37,40,43,46,50,53,56,59,63,66,69,72,76,79,82};


        //for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
        for (int rowNum = startRow; rowNum < rowArray.length; rowNum++) {
            Row row1 = sheet1.getRow(rowArray[rowNum]-1);
            Row row2 = sheet2.getRow(rowArray[rowNum]-1);

            //if (row == null) continue; // Пропустить пустые строки
            for (int colNum = startCol; colNum <= endCol; colNum++) {
                Cell cell1 = row1.getCell(colNum);
                Cell cell2 = row2.getCell(colNum);
                EnhancedCell enhancedCell1 = new EnhancedCell(cell1);
                EnhancedCell enhancedCell2 = new EnhancedCell(cell2);

                System.out.print("Gr. "+enhancedCell1.getValueAsString() + " ");
                System.out.print("Th. "+enhancedCell2.getValueAsString() + " ");
                enhancedCell1.isEmpty();
                enhancedCell2.isEmpty();
            }
            System.out.println("");
        }
        //Iterate through each rows one by one
        //Iterator<Row> rowIterator = sheet1.iterator();
/*        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            //For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                EnhancedCell enhancedCell = new EnhancedCell(cell);
                System.out.print(enhancedCell.getValueAsString() + " ");
            }
            System.out.println("");
        }*/
        file1.close();
    }

}

