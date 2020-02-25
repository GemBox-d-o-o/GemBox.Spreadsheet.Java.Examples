import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Data Types");

        ws.getCell(0, 0).setValue("Cell value examples:");

        // Column width of 25 and 40 characters.
        ws.getColumn(0).setWidth(25 * 256);
        ws.getColumn(1).setWidth(40 * 256);

        // Print gridlines
        ws.getPrintOptions().setPrintGridlines(true);

        int row = 1;

        ws.getCell(++row, 0).setValue("Type");
        ws.getCell(row, 1).setValue("Value");

        ws.getCell(++row, 0).setValue("java.lang.Byte:");
        ws.getCell(row, 1).setValue(Byte.MAX_VALUE);

        ws.getCell(++row, 0).setValue("java.lang.Short:");
        ws.getCell(row, 1).setValue(Short.MAX_VALUE);

        ws.getCell(++row, 0).setValue("java.lang.Long:");
        ws.getCell(row, 1).setValue(Long.MIN_VALUE);

        ws.getCell(++row, 0).setValue("java.lang.Integer:");
        ws.getCell(row, 1).setValue(-5678);

        ws.getCell(++row, 0).setValue("java.lang.Float:");
        ws.getCell(row, 1).setValue(Float.MAX_VALUE);

        ws.getCell(++row, 0).setValue("java.lang.Double:");
        ws.getCell(row, 1).setValue(Double.MAX_VALUE);

        ws.getCell(++row, 0).setValue("java.lang.Boolean:");
        ws.getCell(row, 1).setValue(true);

        ws.getCell(++row, 0).setValue("java.lang.Character:");
        ws.getCell(row, 1).setValue('a');

        ws.getCell(++row, 0).setValue("java.lang.StringBuilder:");
        ws.getCell(row, 1).setValue(new StringBuilder("StringBuilder text."));

        ws.getCell(++row, 0).setValue("java.math.Decimal:");
        ws.getCell(row, 1).setValue(new java.math.BigDecimal(50000));

        ws.getCell(++row, 0).setValue("java.time.LocalDateTime:");
        ws.getCell(row, 1).setValue(java.time.LocalDateTime.now());

        ws.getCell(++row, 0).setValue("System.String:");
        ws.getCell(row++, 1).setValue("Microsoft Excel is a spreadsheet program written and distributed by Microsoft for computers using the Microsoft Windows operating system and Apple Macintosh computers. It is overwhelmingly the dominant spreadsheet application available for these platforms and has been so since version 5 1993 and its bundling as part of Microsoft Office.\n"
                + "Microsoft originally marketed a spreadsheet program called Multiplan in 1982, which was very popular on CP/M systems, but on MS-DOS systems it lost popularity to Lotus 1-2-3. This promoted development of a new spreadsheet called Excel which started with the intention to, in the words of Doug Klunder, 'do everything 1-2-3 does and do it better' . The first version of Excel was released for the Mac in 1985 and the first Windows version (numbered 2.0 to line-up with the Mac and bundled with a run-time Windows environment) was released in November 1987. Lotus was slow to bring 1-2-3 to Windows and by 1988 Excel had started to outsell 1-2-3 and helped Microsoft achieve the position of leading PC software developer. This accomplishment, dethroning the king of the software world, solidified Microsoft as a valid competitor and showed its future of developing graphical software. Microsoft pushed its advantage with regular new releases, every two years or so. The current version is Excel 11, also called Microsoft Office Excel 2003.\n"
                + "Early in its life Excel became the target of a trademark lawsuit by another company already selling a software package named \"Excel.\" As the result of the dispute Microsoft was required to refer to the program as \"Microsoft Excel\" in all of its formal press releases and legal documents. However, over time this practice has slipped.\n"
                + "Excel offers a large number of user interface tweaks, however the essence of UI remains the same as in the original spreadsheet, VisiCalc: the cells are organized in rows and columns, and contain data or formulas with relative or absolute references to other cells.\n"
                + "Excel was the first spreadsheet that allowed the user to define the appearance of spreadsheets (fonts, character attributes and cell appearance). It also introduced intelligent cell recomputation, where only cells dependent on the cell being modified are updated, while previously spreadsheets recomputed everything all the time or waited for a specific user command. Excel has extensive graphing capabilities.\n"
                + "When first bundled into Microsoft Office in 1993, Microsoft Word and Microsoft PowerPoint had their GUIs redesigned for consistency with Excel, the killer app on the PC at the time.\n"
                + "Since 1993 Excel includes support for Visual Basic for Applications (VBA) as a scripting language. VBA is a powerful tool that makes Excel a complete programming environment. VBA and macro recording allow automating routines that otherwise take several manual steps. VBA allows creating forms to handle user input. Automation functionality of VBA exposed Excel as a target for macro viruses.\n"
                + "Excel versions from 5.0 to 9.0 contain various Easter eggs.\n\nFor more information see: http://en.wikipedia.org/wiki/Microsoft_Excel");

        ef.save("Data Types.xlsx");
    }
}