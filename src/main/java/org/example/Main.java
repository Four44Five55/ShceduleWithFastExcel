package org.example;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

public class Main {
    public static void main(String[] args)  {

        //var f = new File("/testExcel/924.xlsx");

        /*try (FileInputStream inputStream = new FileInputStream(new File("/testExcel/924.xlsx"));
             ReadableWorkbook wb = new ReadableWorkbook(inputStream)) {

            for (Sheet sheet : wb.getSheets().toList()) {
                System.out.println("Sheet: " + sheet.getName());
                try (java.util.stream.Stream<Row> rows = sheet.openStream()) {
                    rows.forEach(row -> {
                        row.stream().forEach(cell -> {
                            String value = cell.asString(); // Получаем значение ячейки как строку
                            System.out.print(value + "\t");
                        });
                        System.out.println();
                    });
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }*/

// Укажите путь к вашему Excel файлу
        Path filePath = Paths.get("your_excel_file.xlsx");

        try (Workbook workbook = new Workbook(filePath)) {
            // Получаем лист по имени или индексу (начинается с 0)
            Worksheet worksheet = workbook.getSheet("YourSheetName");  // Или workbook.getSheet(0)


            Sheet sheet = workbook.getFirstSheet();

            if (worksheet != null) {
                // Проходим по всем строкам и колонкам в рабочем листе
                for (int row = 0; row <= worksheet.getLastRowNum(); row++) {
                    for (int col = 0; col <= worksheet.getLastColNum(row); col++) {
                        // Получаем значение ячейки
                        String value = worksheet.сell(row, col).getValue();
                        System.out.print(value != null ? value : "None");
                        System.out.print('\t');  // Табуляция для разделения значений
                    }
                    System.out.println();  // Переход на новую строку после каждой строки ячеек
                }
            } else {
                System.out.println("Рабочий лист не найден");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }



    }
}