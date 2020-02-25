import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;
import java.util.Map;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "TemplateUse.xlsx");

        // Add Sheet
        ExcelWorksheet ws = ef.getWorksheets().insertEmpty(0, "Document Properties");
        ef.getWorksheets().setActiveWorksheet(ws);

        int rowIndex = 0;
        // Read Built-in Document Properties
        ws.getCell(rowIndex++, 0).setValue("Built-in document properties");

        ws.getCell(rowIndex, 0).setValue("Property");
        ws.getCell(rowIndex++, 1).setValue("Value");

        for (Map.Entry<BuiltInDocumentProperties, String> keyValue : ef.getDocumentProperties().getBuiltIn().entrySet()) {
            ws.getCell(rowIndex, 0).setValue(keyValue.getKey().toString());
            ws.getCell(rowIndex++, 1).setValue(keyValue.getValue());
        }

        // Read Custom Document Properties
        ws.getCell(++rowIndex, 0).setValue("Custom Document Properties");

        ws.getCell(++rowIndex, 0).setValue("Property");
        ws.getCell(rowIndex++, 1).setValue("Value");

        // Custom document properties are not supported in XLS
        for (Map.Entry<String, Object> keyValue : ef.getDocumentProperties().getCustom().entrySet()) {
            ws.getCell(rowIndex, 0).setValue(keyValue.getKey());
            ws.getCell(rowIndex++, 1).setValue(keyValue.getValue().toString());
        }

        // Write/Modify Document Properties
        ef.getDocumentProperties().setBuiltIn(BuiltInDocumentProperties.AUTHOR, "John Doe");
        ef.getDocumentProperties().setBuiltIn(BuiltInDocumentProperties.TITLE, "Generated title");

        ws.getColumn(0).setWidth(200, LengthUnit.PIXEL);
        ws.getColumn(1).setWidth(200, LengthUnit.PIXEL);

        ef.save("Document Properties.xlsx");
    }
}