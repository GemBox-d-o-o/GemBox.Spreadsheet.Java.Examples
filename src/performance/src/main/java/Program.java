import com.gembox.spreadsheet.*;

import java.util.concurrent.TimeUnit;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        // If sample exceeds Free version limitations then continue as trial version:
        // https://www.gemboxsoftware.com/spreadsheet-java/help/html/Evaluation_and_Licensing.htm
        SpreadsheetInfo.addFreeLimitReachedListener((event) -> event.setFreeLimitReachedAction(FreeLimitReachedAction.CONTINUE_AS_TRIAL));

        int rowCount = 50000;
        int columnCount = 10;
        String fileFormat = "XLSX";

        System.out.println("Performance sample:");
        System.out.println();
        System.out.println("Row count: " + rowCount);
        System.out.println("Column count: " + columnCount);
        System.out.println("File format: " + fileFormat);
        System.out.println();

        long start = System.nanoTime();

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Performance");

        for (int row = 0; row < rowCount; row++)
            for (int column = 0; column < columnCount; column++)
                ws.getCell(row, column).setValue(row + "_" + column);

        long elapsed = TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - start);
        System.out.println("Generate file (seconds): " + elapsed / 1000.0);

        start = System.nanoTime();

        int cellsCount = 0;
        for (ExcelRow row : ws.getRows())
            for (ExcelCell cell : row.getAllocatedCells())
                cellsCount++;

        elapsed = TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - start);
        System.out.println("Iterate through " + cellsCount + " cells (seconds): " + elapsed / 1000.0);

        start = System.nanoTime();

        ef.save("Report." + fileFormat.toLowerCase());

        elapsed = TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - start);

        System.out.println("Save file (seconds): " + elapsed / 1000.0);
    }
}