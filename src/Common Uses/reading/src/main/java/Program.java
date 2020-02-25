import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "SimpleTemplate.xlsx");

        StringBuilder sb = new StringBuilder();

        // Iterate through all worksheets in an Excel workbook.
        for (ExcelWorksheet sheet : ef.getWorksheets()) {
            sb.append('\n');
            sb.append(String.format("%1$s %2$s %1$s", "-----", sheet.getName()));

            // Iterate through all rows in an Excel worksheet.
            for (ExcelRow row : sheet.getRows()) {
                sb.append('\n');

                // Iterate through all allocated cells in an Excel row.
                for (ExcelCell cell : row.getAllocatedCells()) {
                    if (cell.getValueType() != CellValueType.NULL)
                        sb.append(String.format("%1$s [%2$s]", cell.getValue(), cell.getValueType()));
                    else
                        sb.append("     ");
                }
            }
        }

        System.out.println(sb.toString());
    }
}
