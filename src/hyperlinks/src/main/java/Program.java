import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Hyperlinks");

        ws.getCell("A1").setValue("Hyperlink examples:");

        ExcelCell cell = ws.getCell("B3");
        cell.setValue("GemBoxSoftware");
        cell.getStyle().getFont().setUnderlineStyle(UnderlineStyle.SINGLE);
        cell.getStyle().getFont().setColor(SpreadsheetColor.fromName(ColorName.BLUE));
        cell.getHyperlink().setLocation("https://www.gemboxsoftware.com");
        cell.getHyperlink().setExternal(true);

        cell = ws.getCell("B5");
        cell.setValue("Jump");
        cell.getStyle().getFont().setUnderlineStyle(UnderlineStyle.SINGLE);
        cell.getStyle().getFont().setColor(SpreadsheetColor.fromName(ColorName.BLUE));
        cell.getHyperlink().setToolTip("This is tool tip! This hyperlink jumps to E1.");
        cell.getHyperlink().setLocation(ws.getName() + "!E1");

        ws.getCell("E1").setValue("Destination");

        cell = ws.getCell("B8");
        cell.setFormula("=HYPERLINK(\"https://www.gemboxsoftware.com/spreadsheet-java/examples/excel-cell-hyperlinks/207\", \"Example of HYPERLINK formula\")");
        cell.getStyle().getFont().setUnderlineStyle(UnderlineStyle.SINGLE);
        cell.getStyle().getFont().setColor(SpreadsheetColor.fromName(ColorName.BLUE));

        ef.save("Hyperlinks.xlsx");
    }
}