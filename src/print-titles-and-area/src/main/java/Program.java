import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        // Always print 1st row.
        ExcelWorksheet ws1 = ef.addWorksheet("Sheet1");
        ws1.getNamedRanges().setPrintTitles(ws1.getRow(0), 1);

        // Set print area (from A1 to I120):
        ws1.getNamedRanges().setPrintArea(ws1.getCells().getSubrange("A1", "I120"));

        // Always print columns from A to F.
        ExcelWorksheet ws2 = ef.addWorksheet("Sheet2");
        ws2.getNamedRanges().setPrintTitles(ws2.getColumn(0), 6);

        // Always print columns from A to F and first row.
        ExcelWorksheet ws3 = ef.addWorksheet("Sheet3");
        ws3.getNamedRanges().setPrintTitles(ws3.getRow(0), 1, ws3.getColumn(0), 6);

        // Fill Sheet1 with some data
        for (int i = 0; i < 9; i++)
            ws1.getCell(0, i).setValue("Column " + ExcelColumnCollection.columnIndexToName(i));

        for (int i = 1; i < 120; i++)
            for (int j = 0; j < 9; j++)
                ws1.getCell(i, j).setValue(i + j);

        ef.save("Print Titles and Area.xlsx");
    }
}