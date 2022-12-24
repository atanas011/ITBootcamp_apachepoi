package poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class CreateExcel {

    public static void main(String[] args) {

        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("nums");
        for (int i = 0; i < 5; i++) {
            Row row = sh.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(i + 1);
        }

        try {
            OutputStream os = new FileOutputStream("output.xlsx");
            wb.write(os);
            wb.close();
        } catch (IOException e) {
            System.out.print("Error");
            e.printStackTrace();
        }
    }
}
