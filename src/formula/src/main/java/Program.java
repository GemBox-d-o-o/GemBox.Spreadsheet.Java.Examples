import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Formula");

        int rowIndex = 0;

        ws.getColumn(0).setWidth(35 * 256);
        ws.getColumn(1).setWidth(15 * 256);
        ws.getColumn(2).setWidth(15 * 256);

        ws.getCell(rowIndex++, 0).setValue("Examples of typical formulas usage:");

        ws.getCell(++rowIndex, 0).setValue("Some data:");
        ws.getCell(rowIndex, 1).setValue(3);
        ws.getCell(rowIndex, 2).setValue(4.1);
        ws.getCell(++rowIndex, 1).setValue(5.2);
        ws.getCell(rowIndex, 2).setValue(6);
        ws.getCell(++rowIndex, 1).setValue(7);
        ws.getCell(rowIndex++, 2).setValue(8.3);

        // Named ranges.
        String namedRange = "Range1";
        ws.addNamedRange(namedRange, ws.getCells().getSubrange("B3", "C4"));

        // Floats without first digit.
        ws.getCell(++rowIndex, 0).setValue("Float number without first digit:");
        ws.getCell(rowIndex, 1).setFormula("=.5/23+.1-2");

        // Function using named range.
        ws.getCell(++rowIndex, 0).setValue("Named range:");
        ws.getCell(rowIndex, 1).setFormula("=SUM(" + namedRange + ")");

        // Function's miss argument.
        ws.getCell(++rowIndex, 0).setValue("Function's miss arguments:");
        ws.getCell(rowIndex, 1).setFormula("=Count(1,  ,  ,,,2, 23,,,,,, 34,,,54,,,,  ,)");

        // Functions are case-insensitive.
        ws.getCell(++rowIndex, 0).setValue("Functions are case-insensitive:");
        ws.getCell(rowIndex, 1).setFormula("=cOs( 1 )");

        // Functions.
        ws.getCell(++rowIndex, 0).setValue("Supported functions:");

        String nextFunction;
        ws.getCell(++rowIndex, 0).setValue("Results");
        ws.getCell(rowIndex++, 1).setValue("Formulas");

        nextFunction = "=NOW()+123";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=SECOND(12)/23";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=MINUTE(24)-1343/35";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=(HOUR(56)-23/35)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=WEEKDAY(5)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=YEAR(23)-WEEKDAY(5)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=MONTH(3)-2342/235345";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=((DAY(1)))";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=TIME(1,2,3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=DATE(1,2,3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=RAND()";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=TEXT(\"text\", \"$d\")";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=VAR(1,2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=MOD(1,2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=NOT(FALSE)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=OR(FALSE)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=AND(TRUE)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=FALSE()";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=TRUE()";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=VALUE(3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=LEN(\"hello\")";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=MID(\"hello\",1,1)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=ROUND(1,2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=SIGN(-2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=INT(3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=ABS(-3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=LN(2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=EXP(4)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=SQRT(2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=PI()";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=COS(4)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=SIN(3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=MAX(1,2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=MIN(1,2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=AVERAGE(1,2)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=SUM(1,3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=IF(1,2,3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=COUNT(1,2,3)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        nextFunction = "=SUBTOTAL(1,B3:C5)";
        ws.getCell(rowIndex, 0).setFormula(nextFunction);
        ws.getCell(rowIndex++, 1).setValue(nextFunction);

        // Paranthless checks.
        ws.getCell(++rowIndex, 0).setValue("Paranthless:");
        ws.getCell(rowIndex, 1).setFormula("=((12+2343+34545))");

        // Unary operators.
        ws.getCell(++rowIndex, 0).setValue("Unary operators:");
        ws.getCell(rowIndex, 1).setFormula("=B5%");
        ws.getCell(rowIndex, 2).setFormula("=+++B5");

        // Operand tokens, bool.
        ws.getCell(++rowIndex, 0).setValue("Bool values:");
        ws.getCell(rowIndex, 1).setFormula("=TRUE");
        ws.getCell(rowIndex, 2).setFormula("=FALSE");

        // Operand tokens, int.
        ws.getCell(++rowIndex, 0).setValue("Integer values:");
        ws.getCell(rowIndex, 1).setFormula("=1");
        ws.getCell(rowIndex, 2).setFormula("=20");

        // Operand tokens, num.
        ws.getCell(++rowIndex, 0).setValue("Float values:");
        ws.getCell(rowIndex, 1).setFormula("=.4");
        ws.getCell(rowIndex, 2).setFormula("=2235.5132");

        // Operand tokens, str.
        ws.getCell(++rowIndex, 0).setValue("String values:");
        ws.getCell(rowIndex, 1).setFormula("=\"hello world!\"");

        // Operand tokens, error.
        ws.getCell(++rowIndex, 0).setValue("Error values:");
        ws.getCell(rowIndex, 1).setFormula("=#NULL!");
        ws.getCell(rowIndex, 2).setFormula("=#DIV/0!");

        // Binary operators.
        ws.getCell(++rowIndex, 0).setValue("Binary operators:");
        ws.getCell(rowIndex, 1).setFormula("=(1)-(2)+(3/2+34)/2+12232-32-4");

        ef.save("Formula.xlsx");
    }
}