import com.gembox.spreadsheet.*;
import com.gembox.spreadsheet.tables.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Tables");

        // Add some data
        Object[][] data = {
            { "Worker", "Hours", "Price" },
            { "John Doe", 25, 35.0 },
            { "Jane Doe", 27, 35.0 },
            { "Jack White", 18, 32.0 },
            { "George Black", 31, 35.0 }
        };

        for (int i = 0; i < 5; i++)
            for (int j = 0; j < 3; j++)
                ws.getCell(i, j).setValue(data[i][j]);

        // Set column widths
        ws.getColumn(0).setWidth(100, LengthUnit.PIXEL);
        ws.getColumn(1).setWidth(70, LengthUnit.PIXEL);
        ws.getColumn(2).setWidth(70, LengthUnit.PIXEL);
        ws.getColumn(3).setWidth(70, LengthUnit.PIXEL);
        ws.getColumn(2).getStyle().setNumberFormat("\"$\"#,##0.00");
        ws.getColumn(3).getStyle().setNumberFormat("\"$\"#,##0.00");

        // Create table and enable totals row
        Table table = ws.addTable("Table1", "A1:C5", true);
        table.setHasTotalsRow(true);

        // Add new column
        TableColumn column = table.addColumn();
        column.setName("Total");

        // Populate column
        for (ExcelCell cell : column.getDataRange())
            cell.setFormula("=Table1[Hours] * Table1[Price]");

        // Set totals row function for newly added column and calculate it
        column.setTotalsRowFunction(TotalsRowFunction.SUM);
        column.getRange().calculate();

        // Set table style
        table.setBuiltInStyle(BuiltInTableStyleName.TABLE_STYLE_MEDIUM_2);

        ef.save("Tables.xlsx");
    }
}