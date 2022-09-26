package excel_read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReader {
    // methods
    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {  // add different exception incase it occurs
        //ObjClasi oC = new ObjClasi();
        Scanner scan = new Scanner(System.in);  // creates scanner object

        System.out.print("Enter the customer's name: (for ex.\"coresystem\"): ");
        String toSearch = scan.nextLine();
        int searchColumn = 7, searchA = 10, searchB = 11; // K = 10, L = 11;
        ArrayList<Row> results = new ArrayList<Row>();
        ArrayList<Cell> a = new ArrayList<Cell>();
        ArrayList<Cell> b = new ArrayList<Cell>();
        
        // numbers are rows, letters are columns
        DataFormatter dataFormatter = new DataFormatter(); // Methods that format a value in a cell. It returns the string value after formatting a cell.
        Workbook workbook = WorkbookFactory.create(new FileInputStream("Example.xlsx"));
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); // Evaluates formula cells.
        Sheet sheet = workbook.getSheetAt(0);
        scan.close();
        
        for (Row row : sheet) { // iterate over all rows in the sheet
            Cell cellInSearchColumn = row.getCell(searchColumn); // get the cell in search column (H)
            if (cellInSearchColumn != null) { // if that cell is present
                String cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator); // Return string after formatting cell value.
                if (toSearch.equalsIgnoreCase(cellValue)) { // if cell value equals the searched value
                    results.add(row); // add that row to the results
                }
                
            }
        }

        // print the results
        System.out.println("\nFound results: ");

        for (Row row : results) {
            int rowNumber = row.getRowNum()+1;
            System.out.print("Row " + rowNumber + ":\t");
            for (Cell cell : row) {
                String cellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }
        System.out.println("=============================================================");

        for (Row row : sheet) { // iterate over all rows in the sheet
            Cell cellInSearchColumn = row.getCell(searchColumn); // get the cell in search column (H)
            if (cellInSearchColumn != null) { // if that cell is present
                String cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator); // Return string after formatting cell value.
                if (toSearch.equalsIgnoreCase(cellValue)) { // if cell value equals the searched value
                    a.add(row.getCell(searchA));
                }
                
            }
        }

        System.out.print("Cost weight: ");
        for (Cell cell : a) {
            String cellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);
            System.out.print(cellValue + "\t");
        }

        for (Row row : sheet) { // iterate over all rows in the sheet
            Cell cellInSearchColumn = row.getCell(searchColumn); // get the cell in search column (H)
            if (cellInSearchColumn != null) { // if that cell is present
                String cellValue = dataFormatter.formatCellValue(cellInSearchColumn, formulaEvaluator); // Return string after formatting cell value.
                if (toSearch.equalsIgnoreCase(cellValue)) { // if cell value equals the searched value
                    b.add(row.getCell(searchB));
                }
                
            }
        }

        System.out.println();
        System.out.print("Volume weight: ");
        for (Cell cell : b) {
            String cellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);
            System.out.print(cellValue + "\t");
        }

        workbook.close();
    }
}