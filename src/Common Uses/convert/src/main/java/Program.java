import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "ComplexTemplate.xlsx");

        // In order to achieve the conversion of a loaded Excel file to ODS,
        // or to some other Excel format,
        // we just need to save an ExcelFile object to desired output file format.
        ef.save("Convert.ods");
    }
}