import com.gembox.spreadsheet.*;

import java.util.Random;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Sorting");

        Random rnd = new Random();

        ws.getCell(0, 0).setValue("Sorted numbers");
        ws.getCells().getSubrangeAbsolute(0, 0, 0, 1).setMerged(true);
        for (int i = 1; i < 10; i++)
            ws.getCell(i, 0).setValue(rnd.nextInt(99) + 1);

        ws.getCells().getSubrangeAbsolute(1, 0, 10, 0).sort(false).by(0).apply();

        ws.getCell(0, 2).setValue("Sorted strings");
        ws.getCells().getSubrangeAbsolute(0, 2, 0, 3).setMerged(true);
        ws.getCell(1, 2).setValue("John");
        ws.getCell(2, 2).setValue("Jennifer");
        ws.getCell(3, 2).setValue("Toby");
        ws.getCell(4, 2).setValue("Chloe");

        ws.getCells().getSubrangeAbsolute(1, 2, 4, 2).sort(false).by(0).apply();

        ws.getCell(0, 4).setValue("Sorted by column E and after that by column F");
        ws.getCells().getSubrangeAbsolute(0, 4, 0, 8).setMerged(true);
        for (int i = 1; i < 10; i++)
        {
            ws.getCell(i, 4).setValue(rnd.nextInt(3) + 1);
            ws.getCell(i, 5).setValue(rnd.nextInt(10));
        }

        // Sort by column E ascending and then by column F descending.
        // These sort settings will be saved to output XLSX file because they are active (parameter value is true).
        ws.getCells().getSubrangeAbsolute(1, 4, 10, 5).sort(true).by(0).by(1, true).apply();

        ef.save("Sorting.xlsx");
    }
}