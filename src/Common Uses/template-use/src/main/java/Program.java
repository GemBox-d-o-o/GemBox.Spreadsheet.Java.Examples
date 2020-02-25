import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;
import java.time.LocalDateTime;
import java.util.Random;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "Template.xlsx");

        int workingDays = 8;

        LocalDateTime startDate = LocalDateTime.now().plusDays(-workingDays);
        LocalDateTime endDate = LocalDateTime.now();

        ExcelWorksheet ws = ef.getWorksheet(0);

        // Find cells with placeholder text and set their values.
        RowColumn rowColumnPosition;
        if ((rowColumnPosition = ws.getCells().findText("[Company Name]", true, true)) != null)
            ws.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue("ACME Corp");
        if ((rowColumnPosition = ws.getCells().findText("[Company Address]", true, true)) != null)
            ws.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue("240 Old Country Road, Springfield, IL");
        if ((rowColumnPosition = ws.getCells().findText("[Start Date]", true, true)) != null)
            ws.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue(startDate);
        if ((rowColumnPosition = ws.getCells().findText("[End Date]", true, true)) != null)
            ws.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue(endDate);

        // Copy template row.
        int row = 17;
        ws.getRows().insertCopy(row + 1, workingDays - 1, ws.getRow(row));

        // Fill inserted rows with sample data.
        Random random = new Random();
        for (int i = 0; i < workingDays; i++) {
            ExcelRow currentRow = ws.getRow(row + i);
            currentRow.getCell(1).setValue(startDate.plusDays(i));
            currentRow.getCell(2).setValue(random.nextInt(11) + 1);
        }

        // Calculate formulas in worksheet.
        ws.calculate();

        ef.save("Template Use.xlsx");
    }
}