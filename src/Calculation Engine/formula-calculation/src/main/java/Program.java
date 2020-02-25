import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Formula Calculation");

        // Some formatting.
        ExcelRow row = ws.getRow(0);
        row.getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);

        ExcelColumn col = ws.getColumn(0);
        col.setWidth(250, LengthUnit.PIXEL);
        col.getStyle().setHorizontalAlignment(HorizontalAlignmentStyle.LEFT);
        col = ws.getColumn(1);
        col.setWidth(250, LengthUnit.PIXEL);
        col.getStyle().setHorizontalAlignment(HorizontalAlignmentStyle.RIGHT);

        // Use first row for column headers.
        ws.getCell("A1").setValue("Formula");
        ws.getCell("B1").setValue("Calculated value");

        // Enter some Excel formulas as text in first column.
        ws.getCell("A2").setValue("=1 + 1");
        ws.getCell("A3").setValue("=3 * (2 - 8)");
        ws.getCell("A4").setValue("=3 + ABS(B3)");
        ws.getCell("A5").setValue("=B4 > 15");
        ws.getCell("A6").setValue("=IF(B5, \"Hello world\", \"World hello\")");
        ws.getCell("A7").setValue("=B6 & \" example\"");
        ws.getCell("A8").setValue("=CODE(RIGHT(B7))");
        ws.getCell("A9").setValue("=POWER(B8, 3) * 0.45%");
        ws.getCell("A10").setValue("=SIGN(B9)");
        ws.getCell("A11").setValue("=SUM(B2:B10)");

        // Set text from first column as second row cell's formula.
        int rowIndex = 1;
        while (ws.getCell(rowIndex, 0).getValueType() != CellValueType.NULL)
            ws.getCell(rowIndex, 1).setFormula(ws.getCell(rowIndex++, 0).getStringValue());

        // GemBox.Spreadsheet supports single Excel cell calculation, ...
        ws.getCell("B2").calculate();

        // ... Excel worksheet calculation,
        ws.calculate();

        // ... and whole Excel file calculation.
        ws.getParent().calculate();

        ef.save("Formula Calculation.xlsx");
    }
}