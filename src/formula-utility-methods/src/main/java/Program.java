import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Formula Utility Methods");

        // Fill first column with values.
        for (int i = 0; i < 10; i++)
            ws.getCell(i, 0).setValue(i + 1);

        // Cell B1 has formula '=A1*2', B2 '=A2*2', etc.
        for (int i = 0; i < 10; i++)
            ws.getCell(i, 1).setFormula(String.format("=%1$s*2", CellRange.rowColumnToPosition(i, 0)));

        // Cell C1 has formula '=SUM(A1:B1)', C2 '=SUM(A2:B2)', etc.
        for (int i = 0; i < 10; i++)
            ws.getCell(i, 2).setFormula(String.format("=SUM(A%1$s:B%1$s)", ExcelRowCollection.rowIndexToName(i)));

        // Cell A12 contains sum of all values from the first row.
        ws.getCell("A12").setFormula(String.format("=SUM(A1:%1$s1)", ExcelColumnCollection.columnIndexToName(ws.getRow(0).getAllocatedCells().size() - 1)));

        ef.save("Formula Utility Methods.xlsx");
    }
}