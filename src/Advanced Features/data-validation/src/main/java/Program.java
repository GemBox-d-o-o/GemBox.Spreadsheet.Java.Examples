import com.gembox.spreadsheet.*;

import java.time.LocalDateTime;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Data Validation");

        ws.getCell(0, 0).setValue("Data validation examples:");

        ws.getCell(2, 1).setValue("Decimal greater than 3.14 (on entire row 4):");
        DataValidation dataValidation = new DataValidation(ws.getRow(3).getCells());
        dataValidation.setType(DataValidationType.DECIMAL);
        dataValidation.setOperator(DataValidationOperator.GREATER_THAN);
        dataValidation.setFormula1(3.14);
        dataValidation.setInputMessageTitle("Enter a decimal");
        dataValidation.setInputMessage("Decimal should be greater than 3.14.");
        dataValidation.setErrorTitle("Invalid decimal");
        dataValidation.setErrorMessage("Value should be a decimal greater than 3.14.");
        ws.addDataValidation(dataValidation);
        ws.getCells().getSubrange("A4", "J4").setValue(3.15);

        ws.getCell(7, 1).setValue("List from B9 to B12 (on cell C8):");
        ws.getCell(8, 1).setValue("John");
        ws.getCell(9, 1).setValue("Fred");
        ws.getCell(10, 1).setValue("Hans");
        ws.getCell(11, 1).setValue("Ivan");
        dataValidation = new DataValidation(ws, "C8");
        dataValidation.setType(DataValidationType.LIST);
        dataValidation.setFormula1("=B9:B12");
        dataValidation.setInputMessageTitle("Enter a name");
        dataValidation.setInputMessage("Name should be from the list: John, Fred, Hans, Ivan.");
        dataValidation.setErrorStyle(DataValidationErrorStyle.WARNING);
        dataValidation.setErrorTitle("Invalid name");
        dataValidation.setErrorMessage("Value should be a name from the list: John, Fred, Hans, Ivan.");
        ws.addDataValidation(dataValidation);
        ws.getCell("C8").setValue("John");

        ws.getCell(13, 1).setValue("Date between 2011-01-01 and 2011-12-31 (on cell range C14:E15):");
        dataValidation = new DataValidation(ws.getCells().getSubrange("C14", "E15"));
        dataValidation.setType(DataValidationType.DATE);
        dataValidation.setOperator(DataValidationOperator.BETWEEN);
        dataValidation.setFormula1(LocalDateTime.of(2011, 1, 1, 0, 0));
        dataValidation.setFormula2(LocalDateTime.of(2011, 12, 31, 0, 0));
        dataValidation.setInputMessageTitle("Enter a date");
        dataValidation.setInputMessage("Date should be between 2011-01-01 and 2011-12-31.");
        dataValidation.setErrorStyle(DataValidationErrorStyle.INFORMATION);
        dataValidation.setErrorTitle("Invalid date");
        dataValidation.setErrorMessage("Value should be a date between 2011-01-01 and 2011-12-31.");
        ws.addDataValidation(dataValidation);
        ws.getCells().getSubrange("C14", "E15").setValue(LocalDateTime.of(2011, 1, 1, 0, 0));

        // Column width of 8, 55 and 15 characters.
        ws.getColumn(0).setWidth(8 * 256);
        ws.getColumn(1).setWidth(55 * 256);
        ws.getColumn(2).setWidth(15 * 256);
        ws.getColumn(3).setWidth(15 * 256);
        ws.getColumn(4).setWidth(15 * 256);

        ef.save("Data Validation.xlsx");
    }
}