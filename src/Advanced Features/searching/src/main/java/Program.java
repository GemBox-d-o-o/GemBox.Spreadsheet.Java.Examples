import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "SimpleTemplate.xlsx");

        String searchText = "Apollo 13";

        ExcelWorksheet ws = ef.getWorksheet(0);

        StringBuilder sb = new StringBuilder();

        RowColumn objectPosition = ws.getCells().findText(searchText, false, false);

        if (objectPosition == null) {
            sb.append("Can't find text.\n");
        } else {
            sb.append(searchText + " was launched on " + ws.getCell(objectPosition.getRow(), 2).getValue() + ".\n");

            if (ws.getCell(objectPosition.getRow(), 1).getValue() instanceof String) {
                String nationality = (String) ws.getCell(objectPosition.getRow(), 1).getValue();
                String nationalityText = nationality.trim().toLowerCase();

                int nationalityCounter = 0;

                java.util.Iterator<ExcelCell> iterator = ws.getColumn(1).getCells().iterator();
                while (iterator.hasNext()) {
                    ExcelCell cell = iterator.next();
                    if (cell.getValue() instanceof String && ((String) cell.getValue()).trim().toLowerCase().equals(nationalityText))
                        nationalityCounter++;
                }

                sb.append(String.format("There are %1$s entries for %2$s.", nationalityCounter, nationality));
            }
        }

        System.out.println(sb.toString());
    }
}