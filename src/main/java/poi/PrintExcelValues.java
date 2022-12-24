package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PrintExcelValues {

    public static void main(String[] args) {

        File f = new File("output.xlsx");

        try {
            InputStream is = new FileInputStream(f);
            XSSFWorkbook wb = new XSSFWorkbook(is);
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(0);
                System.out.println(
                        "Cell[" + row.getRowNum() + cell.getColumnIndex() + "] value = " + cell
                );
                wb.close();
            }
        } catch (IOException e) {
            System.out.println("File Not Found");
            e.printStackTrace();
        }
    }
}
