import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "ChartTemplate.xlsx");

        int numberOfEmployees = 4;

        ExcelWorksheet ws = ef.getWorksheet(0);

        // Update named ranges 'Names' and 'Salaries' which are used by preserved chart.
        ws.getNamedRange("Names").setRange(ws.getCells().getSubrangeAbsolute(1, 0, numberOfEmployees, 0));
        ws.getNamedRange("Salaries").setRange(ws.getCells().getSubrangeAbsolute(1, 1, numberOfEmployees, 1));

        // Add data which is used by preserved chart through named ranges 'Names' and 'Salaries'.
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        java.util.Random random = new java.util.Random();
        for (int i = 0; i < numberOfEmployees; i++) {
            ws.getCell(i + 1, 0).setValue(names[i % names.length] + (i < names.length ? "" : " " + (i / names.length + 1)));
            ws.getCell(i + 1, 1).setValue(random.nextInt(4000) + 1000);
        }

        ef.save("Preservation.xlsx");
    }
}