
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class FilesExcel {


    @SuppressWarnings("deprecation")
    public static void writeIntoExcel(String fileName) throws FileNotFoundException, IOException {


        File file = new File(fileName);
        try {
            //проверяем, что если файл не существует то создаем его
            if (!file.exists()) {
                file.createNewFile();
            }

            Workbook book = new HSSFWorkbook();
            Sheet sheet = book.createSheet("TCS SPB");

            // Нумерация начинается с нуля
            Row row = sheet.createRow(0);

            // Мы запишем имя и дату в два столбца
            // имя будет String, а дата рождения --- Date,
            // формата dd.mm.yyyy
            Cell name = row.createCell(0);
            name.setCellValue("Stas");

            Cell birthdate = row.createCell(1);

            DataFormat format = book.createDataFormat();
            CellStyle dateStyle = book.createCellStyle();
            dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
            birthdate.setCellStyle(dateStyle);


            // Нумерация лет начинается с 1900-го
            birthdate.setCellValue(new Date(110, 10, 10));

            // Меняем размер столбца
            sheet.autoSizeColumn(1);

            // Записываем всё в файл
            book.write(new FileOutputStream(fileName));
            book.close();
        } catch(IOException e) {
            throw new RuntimeException(e);
        }
    }

}
