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

        FileInputStream file = new FileInputStream(new File("/testExcel/924.xlsx"));

        //Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        //Get first/desired sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);

        //Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();

        int startRow = 0;
        int endRow = 2;
        int startCol = 0;
        int endCol = 5;


        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            //if (row == null) continue; // Пропустить пустые строки
            for (int colNum = startCol; colNum <= endCol; colNum++) {
                Cell cell = row.getCell(colNum);
                EnhancedCell enhancedCell = new EnhancedCell(cell);
                enhancedCell.isEmpty();
            }
            System.out.println("");
        }

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
        file.close();
    }

}

