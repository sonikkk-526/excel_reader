package excel_read;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

    public static void main(String[] args) {
        //FilePermission permission = new FilePermission(desktop, FilePermission.WRITE);

        // File nF = new File(desktop);

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("FirstExcelSheet"); // first sheet create
        XSSFRow row = sheet.createRow(0); // first row create - A
        XSSFCell cell = row.createCell(0); // first cell create - A-1        
        cell.setCellValue("Testing.."); // give data into A-1 cell
        XSSFRow row1 = sheet.createRow(1);
        XSSFCell cell1 = row.createCell(1);
        cell1.setCellValue("Hello");
        XSSFRow row2 = sheet.createRow(2);
        XSSFCell cell2 = row.createCell(2);
        cell2.setCellValue("World");
        XSSFRow row3 = sheet.createRow(3);
        XSSFCell cell3 = row.createCell(3);
        cell3.setCellValue("Beta.1.0");

        // Output as an excel file
        try (FileOutputStream outputStream = new FileOutputStream("Book1.xlsx")) {
            workbook.write(outputStream);
        } catch (Exception e) {
            System.out.println("An error has occurred during the file output phase, please retry.");
        }

    }
}