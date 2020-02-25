import com.gembox.spreadsheet.*;
import com.gembox.spreadsheet.conditionalformatting.*;
import java.time.LocalDateTime;
import java.util.Random;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Conditional Formatting");

        int rowCount = 20;

        // Specify sheet formatting.
        ws.getRow(0).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        ws.getColumn(0).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(1).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(2).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(3).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(3).getStyle().setNumberFormat("[$$-409]#,##0.00");
        ws.getColumn(4).setWidth(3, LengthUnit.CENTIMETER);
        ws.getColumn(4).getStyle().setNumberFormat("yyyy-mm-dd");

        CellRange cells = ws.getCells();

        // Specify header row.
        cells.get(0, 0).setValue("Departments");
        cells.get(0, 1).setValue("Names");
        cells.get(0, 2).setValue("Years of Service");
        cells.get(0, 3).setValue("Salaries");
        cells.get(0, 4).setValue("Deadlines");

        // Insert random data to sheet.
        Random random = new Random();
        String[] departments = new String[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
        String[] names = new String[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        for (int i = 0; i < rowCount; ++i)
        {
            cells.get(i + 1, 0).setValue(departments[random.nextInt(departments.length)]);
            cells.get(i + 1, 1).setValue(names[random.nextInt(names.length)] + ' ' + (i + 1));
            cells.get(i + 1, 2).setValue(random.nextInt(30) + 1);
            cells.get(i + 1, 3).setValue((random.nextInt(91) + 10) * 100);
            cells.get(i + 1, 4).setValue(LocalDateTime.now().plusDays(random.nextInt(3) - 1));
        }

        // Apply shading to alternate rows in a worksheet using 'Formula' based conditional formatting.
        ws.getConditionalFormatting().addFormula(ws.getCells().getName(), "MOD(ROW(),2)=0").
                getStyle().getFillPattern().setPatternBackgroundColor(SpreadsheetColor.fromName(ColorName.ACCENT_1_LIGHTER_40_PCT));
        ws.getConditionalFormatting().addFormula(ws.getCells().getName(), "MOD(ROW(),2)=1").
                getStyle().getFillPattern().setPatternBackgroundColor(SpreadsheetColor.fromName(ColorName.ACCENT_5_LIGHTER_80_PCT));

        // Apply '2-Color Scale' conditional formatting to 'Years of Service' column. Supported only in XLSX
        ws.getConditionalFormatting().add2ColorScale("C2:C" + (rowCount + 1));

        // Apply '3-Color Scale' conditional formatting to 'Salaries' column. Supported only in XLSX
        ws.getConditionalFormatting().add3ColorScale("D2:D" + (rowCount + 1));

        // Apply 'Data Bar' conditional formatting to 'Salaries' column. Supported only in XLSX
        ws.getConditionalFormatting().addDataBar("D2:D" + (rowCount + 1));

        // Apply 'Icon Set' conditional formatting to 'Years of Service' column. Supported only in XLSX
        ws.getConditionalFormatting().addIconSet("C2:C" + (rowCount + 1)).setIconStyle(SpreadsheetIconStyle.FOUR_TRAFFIC_LIGHTS);

        // Apply green font color to cells in a 'Years of Service' column which have values between 15 and 20.
        ws.getConditionalFormatting().addContainValue("C2:C" + (rowCount + 1), ContainValueOperator.BETWEEN, 15, 20)
                .getStyle().getFont().setColor(SpreadsheetColor.fromName(ColorName.GREEN));

        // Apply double red border to cells in a 'Names' column which contain text 'Doe'.
        ws.getConditionalFormatting().addContainText("B2:B" + (rowCount + 1), ContainTextOperator.CONTAINS, "Doe")
                .getStyle().getBorders().setBorders(MultipleBorders.outside(), SpreadsheetColor.fromName(ColorName.RED), LineStyle.DOUBLE);

        // Apply red shading to cells in a 'Deadlines' column which are equal to yesterday's date. Supported only in XLSX
        ws.getConditionalFormatting().addContainDate("E2:E" + (rowCount + 1), ContainDateOperator.YESTERDAY)
                .getStyle().getFillPattern().setPatternBackgroundColor(SpreadsheetColor.fromName(ColorName.RED));

        // Apply bold font weight to cells in a 'Salaries' column which have top 10 values. Supported only in XLSX
        ws.getConditionalFormatting().addTopOrBottomRanked("D2:D" + (rowCount + 1), false, 10)
                .getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);

        // Apply double underline to cells in a 'Years of Service' column which have below average value. Supported only in XLSX
        ws.getConditionalFormatting().addAboveOrBelowAverage("C2:C" + (rowCount + 1), true)
                .getStyle().getFont().setUnderlineStyle(UnderlineStyle.DOUBLE);

        // Apply italic font style to cells in a 'Departments' column which have duplicate values.
        ws.getConditionalFormatting().addUniqueOrDuplicate("A2:A" + (rowCount + 1), true)
                .getStyle().getFont().setItalic(true);

        ef.save("Conditional Formatting.xlsx");
    }
}