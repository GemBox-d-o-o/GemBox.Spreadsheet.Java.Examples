import com.gembox.spreadsheet.*;
import com.gembox.spreadsheet.charts.*;

import java.time.LocalDateTime;
import java.util.Optional;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        int numberOfEmployees = 4;
        int numberOfYears = 4;

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Chart");

        // Create chart and select data for it.
        ColumnChart chart = (ColumnChart) ws.getCharts().add(ChartType.COLUMN, "B7", "O27");
        chart.selectData(ws.getCells().getSubrangeAbsolute(0, 0, numberOfEmployees, numberOfYears));

        // Set chart title.
        chart.getTitle().setText("Clustered Column Chart");

        // Set axis titles.
        chart.getAxes().getHorizontal().getTitle().setText("Years");
        chart.getAxes().getVertical().getTitle().setText("Salaries");

        // For all charts (except Pie and Bar) value axis is vertical.
        ValueAxis valueAxis = chart.getAxes().getVerticalValue();

        // Set value axis scaling, units, gridlines and tick marks.
        valueAxis.setMinimum(Optional.of(0.0));
        valueAxis.setMaximum(Optional.of(6000.0));
        valueAxis.setMajorUnit(Optional.of(1000.0));
        valueAxis.setMinorUnit(Optional.of(500.0));
        valueAxis.getMajorGridlines().setVisible(true);
        valueAxis.getMinorGridlines().setVisible(true);
        valueAxis.setMajorTickMarkType(TickMarkType.OUTSIDE);
        valueAxis.setMinorTickMarkType(TickMarkType.CROSS);

        // Add data which is used by the chart.
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        java.util.Random random = new java.util.Random();
        for (int i = 0; i < numberOfEmployees; ++i) {
            ws.getCell(i + 1, 0).setValue(names[i % names.length] + (i < names.length ? "" : " " + (i / names.length + 1)));

            for (int j = 0; j < numberOfYears; ++j)
                ws.getCell(i + 1, j + 1).setValue(random.nextInt(4000) + 1000);
        }

        // Set header row and formatting.
        ws.getCell(0, 0).setValue("Name");
        ws.getCell(0, 0).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws.getColumn(0).setWidth((int) LengthUnitConverter.convert(3, LengthUnit.CENTIMETER, LengthUnit.ZERO_CHARACTER_WIDTH_256_TH_PART));
        for (int i = 0, startYear = LocalDateTime.now().getYear() - numberOfYears; i < numberOfYears; ++i) {
            ws.getCell(0, i + 1).setValue(startYear + i);
            ws.getCell(0, i + 1).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
            ws.getCell(0, i + 1).getStyle().setNumberFormat("General");
            ws.getColumn(i + 1).getStyle().setNumberFormat("\"$\"#,##0");
        }

        // Make entire sheet print horizontally centered on a single page with headings and gridlines.
        ExcelPrintOptions printOptions = ws.getPrintOptions();
        printOptions.setPrintHeadings(true);
        printOptions.setPrintHeadings(true);
        printOptions.setHorizontalCentered(true);
        printOptions.setFitWorksheetHeightToPages(1);
        printOptions.setFitWorksheetWidthToPages(1);

        ef.save("Chart Components.xlsx");
    }
}
