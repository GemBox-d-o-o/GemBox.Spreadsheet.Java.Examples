import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "HtmlExport.xlsx");

        ExcelWorksheet ws = ef.getWorksheet(0);

        // Some of the properties from ExcelPrintOptions class are supported in HTML export.
        ws.getPrintOptions().setPrintHeadings(true);
        ws.getPrintOptions().setPrintGridlines(true);

        // Print area can be used to specify custom cell range which should be exported to HTML.
        ws.getNamedRanges().setPrintArea(ws.getCells().getSubrange("A1", "I42"));

        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setHtmlType(HtmlType.HTML);
        options.setSelectionType(SelectionType.ENTIRE_FILE);

        ef.save("HtmlExport.html", options);
    }
}