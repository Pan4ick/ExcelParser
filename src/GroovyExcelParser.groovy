@Grapes([
        @Grab(group = 'org.apache.poi', module = 'poi', version = '4.1.2'),
        @Grab(group = 'org.apache.poi', module = 'poi-ooxml', version = '4.1.2'),
])
import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.ss.util.*
import org.apache.poi.ss.usermodel.*
import java.io.*


class GroovyExcelParser {
    //http://poi.apache.org/spreadsheet/quick-guide.html#Iterator

    static void main(String[] args) {
        FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\spano\\Projects\\ExcelParser\\src\\file.xlsx"));
        // Get the workbook instance for XLS file
        Workbook workbook = WorkbookFactory.create(inputStream);
        // Get first sheet from the workbook
        Sheet sheet = workbook.getSheetAt(0);
        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            // Get iterator to all cells of current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                // Change to getCellType() if using POI 4.x
                CellType cellType = cell.getCellType();
                switch (cellType) {
                    case "_NONE":
                        System.out.print("");
                        System.out.print("\t");
                        break;
                    case "BOOLEAN":
                        System.out.print(cell.getBooleanCellValue());
                        System.out.print("\t");
                        break;
                    case "BLANK":
                        System.out.print("");
                        System.out.print("\t");
                        break;
                    case "FORMULA":
                        // Formula
                        //System.out.print(cell.getCellFormula());
                        System.out.print("\t");

                        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        // Print out value evaluated by formula
                        System.out.print(evaluator.evaluate(cell).getNumberValue());
                        break;
                    case "NUMERIC":
                        System.out.print(cell.getNumericCellValue());
                        System.out.print("\t");
                        break;
                    case "STRING":
                        System.out.print(cell.getStringCellValue());
                        System.out.print("\t");
                        break;
                    case "ERROR":
                        System.out.print("!");
                        System.out.print("\t");
                        break;
                }
            }
            System.out.println("");
        }
    }
}