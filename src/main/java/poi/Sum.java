package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Napisati program koji racuna sumu brojeva iz prvog sheet-a excel tabele koji se zove "nums".
 * Svi brojevi su u prvoj koloni.
 * Program cita red po red iz tabele i upisane brojeve dodaje na sumu.
 * Ukupnu sumu na kraju ispisuje na standardnom izlazu.
 * Program radi i ako se u tabelu doda jos brojeva.
 */

public class Sum {

    public static void main(String[] args) {

        File f = new File("output.xlsx");
        int sum = 0;

        try {
            InputStream in = new FileInputStream(f);
            XSSFWorkbook wb = new XSSFWorkbook(in);
            // I
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                sum += sheet.getRow(i).getCell(0).getNumericCellValue();
            }
//            // II
//            Sheet sheet = wb.getSheet("nums");
//            int count = 1;
//            Row row = sheet.getRow(0);
//            Cell cell = row.getCell(0);
//            while (cell != null && !cell.toString().equals("")) {
//                sum += cell.getNumericCellValue();
//                row = sheet.getRow(count);
//                if (row == null) break;
//                cell = row.getCell(0);
//                count++;
//            }
            System.out.print("Sum = " + sum);
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
