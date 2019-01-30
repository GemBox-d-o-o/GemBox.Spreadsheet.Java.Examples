import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Hello World");

        ws.getCell(0, 0).setValue("English:");
        ws.getCell(0, 1).setValue("Hello");

        ws.getCell(1, 0).setValue("Russian:");
        // Using UNICODE string.
        ws.getCell(1, 1).setValue(new String(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' }));

        ws.getCell(2, 0).setValue("Chinese:");
        // Using UNICODE string.
        ws.getCell(2, 1).setValue(new String(new char[] { '\u4f60', '\u597d' }));

        ws.getCell(4, 0).setValue("In order to see Russian and Chinese characters you need to have appropriate fonts on your PC.");
        ws.getCells().getSubrangeAbsolute(4, 0, 4, 7).setMerged(true);

        ef.save("Hello World.xlsx");
    }
}