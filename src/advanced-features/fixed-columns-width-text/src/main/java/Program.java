import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        // Define columns width (for input file format)
        FixedWidthLoadOptions loadOptions = new FixedWidthLoadOptions(
                new FixedWidthColumn(8),
                new FixedWidthColumn(8),
                new FixedWidthColumn(8));

        // Load file
        ExcelFile ef = ExcelFile.load(resourcesFolder + "FixedColumnsWidthText.prn", loadOptions);

        // Modify file
        ef.getWorksheets().getActiveWorksheet().getUsedCellRange(true).sort(false).by(1).apply();

        // Define columns width (for output file format)
        FixedWidthSaveOptions saveOptions = new FixedWidthSaveOptions(
                new FixedWidthColumn(8),
                new FixedWidthColumn(8),
                new FixedWidthColumn(8));

        ef.save("SortedFixedColumnsWidthText.prn", saveOptions);
    }
}