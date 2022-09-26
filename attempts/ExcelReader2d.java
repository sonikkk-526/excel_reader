/*
 * The goal of this file is to simulate a GUI (as of now).
*/

package excel_reader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader2d {
    ExcelReader link = new ExcelReader();
    Scanner scan = new Scanner(System.in);

    // instance data
    private static int numberGrid[][] = null;
    private static String stringGrid[][] = null;

    // constructor
    public ExcelReader2d(String desLink) {

    }

    // methods
    public void ExeScan() throws FileNotFoundException, IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(link.getImpoFileLink()));
        XSSFSheet sheet = workbook.getSheetAt(0);
        System.out.print("Enter the rows that you want to search for: (for ex. the rows that stores customer's name) ");
        System.out.println("Please noted that if the first row, it will be A, enter '0'. If \"B\", enter 1, so on...");
        int searchRows = scan.nextInt();
        XSSFRow row = sheet.getRow(searchRows);
        try {
            for (int j = 3; j < row.getRowNum(); j++) {
                for (int = i; i < ; i++) {
					
                }
            }
        }
        
        
            
    }

}