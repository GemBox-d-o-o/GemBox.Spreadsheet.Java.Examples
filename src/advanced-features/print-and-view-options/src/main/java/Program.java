import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Print and View Options");

        ws.getCell("M1").setValue("This worksheet shows how to set various print related and view related options.");
        ws.getCell("M2").setValue("To see results of print options, go to Print and Page Setup dialogs in MS Excel.");
        ws.getCell("M3").setValue("Notice that print and view options are worksheet based, not workbook based.");

        // Print options:
        ExcelPrintOptions printOptions = ws.getPrintOptions();
        printOptions.setPrintGridlines(true);
        printOptions.setPrintHeadings(true);
        printOptions.setPortrait(false);
        printOptions.setPaperType(PaperType.A3);
        printOptions.setNumberOfCopies(5);

        // View options:
        ws.getViewOptions().setFirstVisibleColumn(3);
        ws.getViewOptions().setShowColumnsFromRightToLeft(true);
        ws.getViewOptions().setZoom(123);

        // Set print area
        ws.getNamedRanges().setPrintArea(ws.getCells().getSubrange("E1", "U7"));

        ef.save("Print and View Options.xlsx");
    }
}