import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Comments");

        ws.getCell(0, 0).setValue("Comment examples:");

        ws.getCell(2, 1).getComment().setText("Empty cell.");

        ws.getCell(4, 1).setValue(5);
        ws.getCell(4, 1).getComment().setText("Cell with a number.");

        ws.getCell("B7").setValue("Cell B7");

        ExcelComment comment = ws.getCell("B7").getComment();
        comment.setText("Some formatted text.\nComment is:\na) multiline,\nb) large,\nc) visible, and \nd) formatted.");
        comment.setVisible(true);
        comment.setTopLeftCell(new AnchorCell(ws.getColumn(3), ws.getRow(4), true));
        comment.setBottomRightCell(new AnchorCell(ws.getColumn(5), ws.getRow(10), false));

        // Get first 20 characters of a string
        FormattedCharacterRange characters = comment.getCharacters(0, 20);

        // Apply color, italic and size to selected characters
        characters.getFont().setColor(SpreadsheetColor.fromName(ColorName.ORANGE));
        characters.getFont().setItalic(true);
        characters.getFont().setSize(300);

        // Apply color to 'formatted' part of text
        comment.getCharacters(5, 9).getFont().setColor(SpreadsheetColor.fromName(ColorName.RED));

        ef.save("Comments.xlsx");
    }
}