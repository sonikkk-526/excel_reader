package excel_read;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Lab {
    public static void main(String[] args) throws Exception, IOException {
        List<Row> compare = new ArrayList<Row>();

        Workbook workbook = WorkbookFactory.create(new FileInputStream("test.XLSX"));
        DataFormatter dataFormatter = new DataFormatter(); // Methods that format a value in a cell. It returns the string value after formatting a cell.
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator(); // Evaluates formula cells.
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(7);
                    if (cell != null) {
                        String cellValue = dataFormatter.formatCellValue(cell, formulaEvaluator);
                        compare.add(row);
                    }
                } else {
                    continue;
                }
            }
        }

        // Print found result
        //for (int i = 0; i <= compare.length()-1; i++) {

        //}
    }
}