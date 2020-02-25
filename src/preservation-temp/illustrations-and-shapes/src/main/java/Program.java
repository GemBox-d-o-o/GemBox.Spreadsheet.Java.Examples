import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "IllustrationsAndShapes.xlsx");

        ef.save("Illustrations and Shapes.xlsx");
    }
}