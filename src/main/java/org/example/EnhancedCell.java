package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.SimpleDateFormat;
import java.util.Date;

import static org.apache.poi.ss.usermodel.CellType.BLANK;

//import static org.apache.poi.ss.usermodel.CellType.BLANK;

public class EnhancedCell {
    private Cell cell;

    public EnhancedCell(Cell cell) {
        this.cell = cell;
    }

    public boolean getEmpty() {
        if (cell.getCellType() == BLANK) {
            return true;
        } else {
            return false;
        }
    }

    // Добавляем новый метод:  получение значения в виде строки, обрабатывая различные типы данных
    public String getValueAsString() {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // Проверка на наличие даты в ячейке
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Получение даты из ячейки
                    Date date = cell.getDateCellValue();
                    // Форматирование даты
                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yy");
                    // Возврат отформатированной даты
                    return dateFormat.format(date);
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula(); // Или вычисление формулы, если нужно
            case BLANK:
                return "";
            case ERROR:
                return "Ошибка";
            default:
                return "Неизвестный тип ячейки";
        }
    }

}
