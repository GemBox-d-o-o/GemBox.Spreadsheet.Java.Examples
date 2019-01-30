import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Inline Text Formatting");

        ws.getCell(0, 0).setValue("Inline text formatting examples:");
        ws.getPrintOptions().setPrintGridlines(true);

        // Column width of 20 characters.
        ws.getColumn(0).setWidth(20 * 256);

        ws.getCell(2, 0).setValue("This is big and red text!");

        // Apply size to 'big and red' part of text
        ws.getCell(2, 0).getCharacters(8, 11).getFont().setSize(400);

        // Apply color to 'red' part of text
        ws.getCell(2, 0).getCharacters(16, 3).getFont().setColor(SpreadsheetColor.fromName(ColorName.RED));

        // Format cell content
        ws.getCell(4, 0).setValue("Formatting selected characters with GemBox.Spreadsheet for Java component.");
        ws.getCell(4, 0).getStyle().getFont().setColor(SpreadsheetColor.fromName(ColorName.BLUE));
        ws.getCell(4, 0).getStyle().getFont().setItalic(true);
        ws.getCell(4, 0).getStyle().setWrapText(true);

        // Get characters from index 36 to the end of string
        FormattedCharacterRange characters = ws.getCell(4, 0).getCharacters(36);

        // Apply color and underline to selected characters
        characters.getFont().setColor(SpreadsheetColor.fromName(ColorName.ORANGE));
        characters.getFont().setUnderlineStyle(UnderlineStyle.SINGLE);

        // Write selected characters
        ws.getCell(6, 0).setValue("Selected characters: " + characters.getText());

        ef.save("Inline Text Formatting.xlsx");
    }
}