import com.gembox.spreadsheet.*;
import com.gembox.spreadsheet.charts.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Chart");

        int numberOfEmployees = 4;

        // Create Excel chart and select data for it.
        ExcelChart chart = ws.getCharts().add(ChartType.BAR, "D2", "M25");
        chart.selectData(ws.getCells().getSubrangeAbsolute(0, 0, numberOfEmployees, 1), true);

        // Add data which is used by the Excel chart.
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        java.util.Random random = new java.util.Random();
        for (int i = 0; i < numberOfEmployees; i++) {
            ws.getCell(i + 1, 0).setValue(names[i % names.length] + (i < names.length ? "" : " " + (i / names.length + 1)));
            ws.getCell(i + 1, 1).setValue(random.nextInt(4000) + 1000);
        }

        // Set header row and formatting.
        ws.getCell(0, 0).setValue("Name");
        ws.getCell(0, 1).setValue("Salary");
        ws.getCell(0, 1).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws.getCell(0, 0).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws.getColumn(0).setWidth((int) LengthUnitConverter.convert(3, LengthUnit.CENTIMETER, LengthUnit.ZERO_CHARACTER_WIDTH_256_TH_PART));
        ws.getColumn(1).getStyle().setNumberFormat("\"$\"#,##0");

        ef.save("Chart.xlsx");
    }
}