import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        String inputPassword = "inpass";
        String outputPassword = "outpass";

        XlsLoadOptions loadOptions = new XlsLoadOptions();
        loadOptions.setPassword(inputPassword);
        ExcelFile ef = ExcelFile.load(resourcesFolder + "XlsEncryption.xls", loadOptions);

        XlsSaveOptions saveOptions = new XlsSaveOptions();
        saveOptions.setPassword(outputPassword);
        ef.save("XLS Encryption.xls", saveOptions);
    }
}