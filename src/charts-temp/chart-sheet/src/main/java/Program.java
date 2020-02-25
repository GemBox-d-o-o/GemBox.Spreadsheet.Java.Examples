import com.gembox.spreadsheet.*;
import com.gembox.spreadsheet.charts.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        int numberOfEmployees = 4;

        ExcelWorksheet ws1 = ef.addWorksheet("SourceSheet");

        // Add data which is used by the Excel chart.
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        java.util.Random random = new java.util.Random();
        for (int i = 0; i < numberOfEmployees; i++)
        {
            ws1.getCell(i + 1, 0).setValue(names[i % names.length] + (i < names.length ? "" : " " + (i / names.length + 1)));
            ws1.getCell(i + 1, 1).setValue(random.nextInt(4000) + 1000);
        }

        // Set header row and formatting.
        ws1.getCell(0, 0).setValue("Name");
        ws1.getCell(0, 1).setValue("Salary");
        ws1.getCell(0, 1).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws1.getCell(0, 0).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws1.getColumn(0).setWidth((int)LengthUnitConverter.convert(3, LengthUnit.CENTIMETER, LengthUnit.ZERO_CHARACTER_WIDTH_256_TH_PART));
        ws1.getColumn(1).getStyle().setNumberFormat("\"$\"#,##0");

        // Create Excel chart sheet.
        ExcelWorksheet ws2 = ef.getWorksheets().add(SheetType.CHART, "ChartSheet");

        // Create Excel chart and select data for it.
        // You cannot set the size of the chart area when the chart is located on a chart sheet, it will snap to maximum size on the chart sheet.
        ExcelChart chart = ws2.getCharts().add(ChartType.BAR, 0, 0, 0, 0, LengthUnit.CENTIMETER);
        chart.selectData(ws1.getCells().getSubrangeAbsolute(0, 0, numberOfEmployees, 1), true);

        ef.save("Chart Sheet.xlsx");
    }
}