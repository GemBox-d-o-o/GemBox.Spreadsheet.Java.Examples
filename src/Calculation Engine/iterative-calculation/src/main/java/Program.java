import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        // Set calculation options.
        ef.getCalculationOptions().setMaximumIterations(10);
        ef.getCalculationOptions().setMaximumChange(0.05);
        ef.getCalculationOptions().setEnableIterativeCalculation(true);

        // Add new worksheet
        ExcelWorksheet ws = ef.addWorksheet("Iterative Calculation");

        // Some column formatting.
        ws.getColumn(0).setWidth(50, LengthUnit.PIXEL);
        ws.getColumn(1).setWidth(100, LengthUnit.PIXEL);

        // Simple example of circular reference limited by MaximumIterations in column A.
        ws.getCell("A1").setFormula("=A2");
        ws.getCell("A2").setFormula("=A1 + 1");

        // Simple example of circular reference limited by MaximumChange in column B.
        ws.getCell("B1").setValue(100000.0);
        ws.getCell("B2").setFormula("=B3 * 0.03");
        ws.getCell("B3").setFormula("=B1 + B2");

        // Calculate all cells.
        ws.calculate();

        ef.save("Iterative Calculation.xlsx");
    }
}