import com.gembox.spreadsheet.*;

import java.time.LocalDateTime;
import java.util.Random;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Filtering");

        int rowCount = 100;

        // Specify sheet formatting.
        ws.getRow(0).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws.getColumn(0).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(1).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(2).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(2).getStyle().setNumberFormat("[$$-409]#,##0.00");
        ws.getColumn(3).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(3).getStyle().setNumberFormat("yyyy-mm-dd");

        CellRange cells = ws.getCells();

        // Specify header row.
        cells.get(0, 0).setValue("Departments");
        cells.get(0, 1).setValue("Names");
        cells.get(0, 2).setValue("Salaries");
        cells.get(0, 3).setValue("Deadlines");

        // Insert random data to sheet.
        Random random = new Random();
        String[] departments = new String[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        for (int i = 0; i < rowCount; ++i) {
            cells.get(i + 1, 0).setValue(departments[random.nextInt(departments.length)]);
            cells.get(i + 1, 1).setValue(names[random.nextInt(names.length)] + ' ' + (i + 1));
            cells.get(i + 1, 2).setValue((random.nextInt(91) + 10) * 100);
            cells.get(i + 1, 3).setValue(LocalDateTime.now().plusDays(random.nextInt(3) - 1));
        }

        // Specify range which will be filtered.
        CellRange filterRange = ws.getCells().getSubrangeAbsolute(0, 0, rowCount, 3);

        // Show only rows which satisfy following conditions:
        // - 'Departments' value is either "Legal" or "Marketing" or "Finance" and
        // - 'Names' value contains letter 'e' and
        // - 'Salaries' value is in the top 20 percent of all 'Salaries' values and
        // - 'Deadlines' value is today's date.
        // Shown rows are then sorted by 'Salaries' values in the descending order.
        filterRange.filter().
                byValues(0, "Legal", "Marketing", "Finance").
                byCustom(1, FilterOperator.EQUAL, "*e*").
                byTop10(2, true, true, 20).
                byDynamic(3, DynamicFilterType.TODAY).
                sortBy(2, true).
                apply();

        ef.save("Filtering.xlsx");
    }
}