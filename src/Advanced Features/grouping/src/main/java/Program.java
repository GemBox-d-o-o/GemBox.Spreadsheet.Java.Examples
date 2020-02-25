import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Grouping");

        ws.getCell(0, 0).setValue("Cell grouping examples:");

        // Vertical grouping.
        ws.getCell(2, 0).setValue("GroupA Start");
        ws.getRow(2).setOutlineLevel(1);
        ws.getCell(3, 0).setValue("A");
        ws.getRow(3).setOutlineLevel(1);
        ws.getCell(4, 1).setValue("GroupB Start");
        ws.getRow(4).setOutlineLevel(2);
        ws.getCell(5, 1).setValue("B");
        ws.getRow(5).setOutlineLevel(2);
        ws.getCell(6, 1).setValue("GroupB End");
        ws.getRow(6).setOutlineLevel(2);
        ws.getCell(7, 0).setValue("GroupA End");
        ws.getRow(7).setOutlineLevel(1);
        // Put outline row buttons above groups.
        ws.getViewOptions().setOutlineRowButtonsBelow(false);

        // Horizontal grouping (collapsed).
        ws.getCell("E2").setValue("Gr.C Start");
        ws.getColumn("E").setOutlineLevel(1);
        ws.getColumn("E").setHidden(true);
        ws.getCell("F2").setValue("C");
        ws.getColumn("F").setOutlineLevel(1);
        ws.getColumn("F").setHidden(true);
        ws.getCell("G2").setValue("Gr.C End");
        ws.getColumn("G").setOutlineLevel(1);
        ws.getColumn("G").setHidden(true);
        ws.getColumn("H").setCollapsed(true);

        ef.save("Grouping.xlsx");
    }
}