import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "NumberFormat.xlsx");

        ExcelWorksheet ws = ef.getWorksheet(0);

        ws.getCell(0, 2).setValue("ExcelCell.Value");
        ws.getColumn(2).getStyle().setNumberFormat("@");

        ws.getCell(0, 3).setValue("CellStyle.NumberFormat");
        ws.getColumn(3).getStyle().setNumberFormat("@");

        ws.getCell(0, 4).setValue("ExcelCell.GetFormattedValue()");
        ws.getColumn(4).getStyle().setNumberFormat("@");

        for (int i = 1; i < ws.getRows().size(); i++) {
            ExcelCell sourceCell = ws.getCell(i, 0);

            ws.getCell(i, 2).setValue(sourceCell.getValue() == null ? null : sourceCell.getValue().toString());
            ws.getCell(i, 3).setValue(sourceCell.getStyle().getNumberFormat());
            ws.getCell(i, 4).setValue(sourceCell.getFormattedValue());
        }

        ef.save("Number Format.xlsx");
    }
}