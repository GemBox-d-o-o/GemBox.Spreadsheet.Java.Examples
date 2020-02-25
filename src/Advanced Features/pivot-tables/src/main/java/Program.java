import com.gembox.spreadsheet.*;
import com.gembox.spreadsheet.pivottables.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        ExcelWorksheet ws1 = ef.addWorksheet("SourceSheet");

        // Specify sheet formatting.
        ws1.getRow(0).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws1.getColumn(0).setWidth(3, LengthUnit.CENTIMETER);
        ws1.getColumn(1).setWidth(3, LengthUnit.CENTIMETER);
        ws1.getColumn(2).setWidth(3, LengthUnit.CENTIMETER);
        ws1.getColumn(3).setWidth(3, LengthUnit.CENTIMETER);
        ws1.getColumn(3).getStyle().setNumberFormat("[$$-409]#,##0.00");

        CellRange cells = ws1.getCells();

        // Specify header row.
        cells.get(0, 0).setValue("Departments");
        cells.get(0, 1).setValue("Names");
        cells.get(0, 2).setValue("Years of Service");
        cells.get(0, 3).setValue("Salaries");

        // Insert random data to sheet.
        java.util.Random random = new java.util.Random();
        String[] departments = new String[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        String[] years = new String[] { "1-10", "11-20", "21-30", "over 30" };
        for (int i = 0; i < 100; ++i) {
            cells.get(i + 1, 0).setValue(departments[random.nextInt(departments.length)]);
            cells.get(i + 1, 1).setValue(names[random.nextInt(names.length)] + ' ' + (i + 1));
            cells.get(i + 1, 2).setValue(years[random.nextInt(years.length)]);
            cells.get(i + 1, 3).setValue((random.nextInt(91) + 10) * 100);
        }

        // Create pivot cache from cell range "SourceSheet!A1:D100".
        PivotCache cache = ef.getPivotCaches().addWorksheetSource("SourceSheet!A1:D100");

        // Create new sheet for pivot table.
        ExcelWorksheet ws2 = ef.addWorksheet("PivotSheet");

        // Create pivot table "Company Profile" using the specified pivot cache and add it to the worksheet at the cell location 'A1'.
        PivotTable table = ws2.addPivotTable(cache, "Company Profile", "A1");

        // Aggregate 'Names' values into count value and show it as a percentage of row.
        PivotField field = table.getDataFields().add("Names");
        field.setFunction(PivotFieldCalculationType.COUNT);
        field.setShowDataAs(PivotFieldDisplayFormat.PERCENTAGE_OF_ROW);
        field.setName("% of Empl.");

        // Aggregate 'Salaries' values into average value.
        field = table.getDataFields().add("Salaries");
        field.setFunction(PivotFieldCalculationType.AVERAGE);
        field.setName("Avg. Salary");
        field.setNumberFormat("[$$-409]#,##0.00");

        // Group rows into 'Departments'.
        table.getRowFields().add("Departments");

        // Group columns first into 'Years of Service' and then into 'Values' (count 'Names' and average 'Salaries').
        table.getColumnFields().add("Years of Service");
        table.getColumnFields().add(table.getDataPivotField());

        // Specify the string to be displayed in row and column header.
        table.setRowHeaderCaption("Departments");
        table.setColumnHeaderCaption("Years of Service");

        // Do not show grand totals for rows.
        table.setRowGrandTotals(false);

        // Set pivot table style.
        table.setBuiltInStyle(BuiltInPivotStyleName.PIVOT_STYLE_MEDIUM_7);

        ef.save("Pivot Tables.xlsx");
    }
}