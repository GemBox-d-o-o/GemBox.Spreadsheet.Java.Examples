import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Cell Referencing");

        ws.getCell(0, 0).setValue("Cell referencing examples:");

        ws.getCell("B2").setValue("Cell B2.");
        ws.getCell(6, 0).setValue("Cell in row 7 and column A.");

        ws.getRow(2).getCell(0).setValue("Cell in row 3 and column A.");
        ws.getRow("4").getCell("B").setValue("Cell in row 4 and column B.");

        ws.getColumn(2).getCell(4).setValue("Cell in column C and row 5.");
        ws.getColumn("AA").getCell("6").setValue("Cell in AA column and row 6.");

        // Referencing Excel row's cell range.
        CellRange cr = ws.getRow(7).getCells();

        cr.get(0).setValue(cr.getIndexingMode().toString());
        cr.get(3).setValue("D8");
        cr.get("B").setValue("B8");

        // Referencing Excel column's cell range.
        cr = ws.getColumn(7).getCells();

        cr.get(0).setValue(cr.getIndexingMode().toString());
        cr.get(2).setValue("H3");
        cr.get("5").setValue("H5");

        // Referencing arbitrary Excel cell range.
        cr = ws.getCells().getSubrange("I2", "L8");
        cr.getStyle().getBorders().setBorders(MultipleBorders.outside(), SpreadsheetColor.fromArgb(0, 0, 128), LineStyle.DASHED);

        cr.get("J7").setValue(cr.getIndexingMode().toString());
        cr.get(0, 0).setValue("I2");
        cr.get("J3").setValue("J3");
        cr.get(4).setValue("I3"); // Cell range width is 4 (I J K L).

        ws.getPrintOptions().setPrintGridlines(true);
        ws.getPrintOptions().setPrintHeadings(true);

        ef.save("Cell Referencing.xlsx");
    }
}