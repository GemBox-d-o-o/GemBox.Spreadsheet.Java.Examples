import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Styles and Formatting");

        ws.getCell(0, 1).setValue("Cell style examples:");
        ws.getPrintOptions().setPrintGridlines(true);

        int row = 0;

        // Column width of 4, 30 and 36 characters.
        ws.getColumn(0).setWidth(4 * 256);
        ws.getColumn(1).setWidth(30 * 256);
        ws.getColumn(2).setWidth(36 * 256);

        ws.getCell(row += 2, 1).setValue(".Style.Borders.SetBorders(...)");
        ws.getCell(row, 2).getStyle().getBorders().setBorders(MultipleBorders.outside(), SpreadsheetColor.fromArgb(252, 1, 1), LineStyle.THIN);

        ws.getCell(row += 2, 1).setValue(".Style.FillPattern.SetPattern(...)");
        ws.getCell(row, 2).getStyle().getFillPattern().setPattern(FillPatternStyle.THIN_HORIZONTAL_CROSSHATCH, SpreadsheetColor.fromName(ColorName.GREEN), SpreadsheetColor.fromName(ColorName.YELLOW));

        ws.getCell(row += 2, 1).setValue(".Style.Font.Color =");
        ws.getCell(row, 2).setValue("Color.Blue");
        ws.getCell(row, 2).getStyle().getFont().setColor(SpreadsheetColor.fromName(ColorName.BLUE));

        ws.getCell(row += 2, 1).setValue(".Style.Font.Italic =");
        ws.getCell(row, 2).setValue("true");
        ws.getCell(row, 2).getStyle().getFont().setItalic(true);

        ws.getCell(row += 2, 1).setValue(".Style.Font.Name =");
        ws.getCell(row, 2).setValue("Comic Sans MS");
        ws.getCell(row, 2).getStyle().getFont().setName("Comic Sans MS");

        ws.getCell(row += 2, 1).setValue(".Style.Font.ScriptPosition =");
        ws.getCell(row, 2).setValue("ScriptPosition.Superscript");
        ws.getCell(row, 2).getStyle().getFont().setScriptPosition(ScriptPosition.SUPERSCRIPT);

        ws.getCell(row += 2, 1).setValue(".Style.Font.Size =");
        ws.getCell(row, 2).setValue("18 * 20");
        ws.getCell(row, 2).getStyle().getFont().setSize(18 * 20);

        ws.getCell(row += 2, 1).setValue(".Style.Font.Strikeout =");
        ws.getCell(row, 2).setValue("true");
        ws.getCell(row, 2).getStyle().getFont().setStrikeout(true);

        ws.getCell(row += 2, 1).setValue(".Style.Font.UnderlineStyle =");
        ws.getCell(row, 2).setValue("UnderlineStyle.Double");
        ws.getCell(row, 2).getStyle().getFont().setUnderlineStyle(UnderlineStyle.DOUBLE);

        ws.getCell(row += 2, 1).setValue(".Style.Font.Weight =");
        ws.getCell(row, 2).setValue("ExcelFont.BoldWeight");
        ws.getCell(row, 2).getStyle().getFont().setWeight(ExcelFont.BOLD_WEIGHT);

        ws.getCell(row += 2, 1).setValue(".Style.HorizontalAlignment =");
        ws.getCell(row, 2).setValue("HorizontalAlignmentStyle.Center");
        ws.getCell(row, 2).getStyle().setHorizontalAlignment(HorizontalAlignmentStyle.CENTER);

        ws.getCell(row += 2, 1).setValue(".Style.Indent");
        ws.getCell(row, 2).setValue("five");
        ws.getCell(row, 2).getStyle().setHorizontalAlignment(HorizontalAlignmentStyle.LEFT);
        ws.getCell(row, 2).getStyle().setIndent(5);

        ws.getCell(row += 2, 1).setValue(".Style.IsTextVertical = ");
        ws.getCell(row, 2).setValue("true");
        // Set row height to 60 points.
        ws.getRow(row).setHeight(60 * 20);
        ws.getCell(row, 2).getStyle().setTextVertical(true);

        ws.getCell(row += 2, 1).setValue(".Style.NumberFormat");
        ws.getCell(row, 2).setValue(1234);
        ws.getCell(row, 2).getStyle().setNumberFormat("#,##0.00 [$Krakozhian Money Units]");

        ws.getCell(row += 2, 1).setValue(".Style.Rotation");
        ws.getCell(row, 2).setValue("35 degrees up");
        ws.getCell(row, 2).getStyle().setRotation(35);

        ws.getCell(row += 2, 1).setValue(".Style.ShrinkToFit");
        ws.getCell(row, 2).setValue("This property is set to true so this text appears shrunk.");
        ws.getCell(row, 2).getStyle().setShrinkToFit(true);

        ws.getCell(row += 2, 1).setValue(".Style.VerticalAlignment =");
        ws.getCell(row, 2).setValue("VerticalAlignmentStyle.Top");
        // Set row height to 30 points.
        ws.getRow(row).setHeight(30 * 20);
        ws.getCell(row, 2).getStyle().setVerticalAlignment(VerticalAlignmentStyle.TOP);

        ws.getCell(row += 2, 1).setValue(".Style.WrapText");
        ws.getCell(row, 2).setValue("This property is set to true so this text appears broken into multiple lines.");
        ws.getCell(row, 2).getStyle().setWrapText(true);

        ef.save("Styles and Formatting.xlsx");
    }
}