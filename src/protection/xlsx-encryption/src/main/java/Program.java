import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        String inputPassword = "inpass";
        String outputPassword = "outpass";

        XlsxLoadOptions loadOptions = new XlsxLoadOptions();
        loadOptions.setPassword(inputPassword);
        ExcelFile ef = ExcelFile.load(resourcesFolder + "XlsxEncryption.xlsx", loadOptions);

        XlsxSaveOptions saveOptions = new XlsxSaveOptions();
        saveOptions.setPassword(outputPassword);
        ef.save("XLSX Encryption.xlsx", saveOptions);
    }
}