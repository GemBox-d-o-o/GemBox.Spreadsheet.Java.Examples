import com.gembox.examples.Util;
import com.gembox.spreadsheet.*;

class Program {

    private static final String resourcesFolder = Util.resourcesFolder();

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.load(resourcesFolder + "IllustrationsAndShapes.xlsx");

        StringBuilder sb = new StringBuilder();

        ExcelWorksheet ws = ef.getWorksheet(0);

        sb.append(String.format("Sheet left margin is: %1$s pixels.", Math.round(LengthUnitConverter.convert(ws.getPrintOptions().getLeftMargin(), LengthUnit.INCH, LengthUnit.PIXEL))));
        sb.append('\n');

        sb.append(String.format("Width of column A is: %1$s pixels.", Math.round(LengthUnitConverter.convert(ws.getColumn(0).getWidth(), LengthUnit.ZERO_CHARACTER_WIDTH_256_TH_PART, LengthUnit.PIXEL))));
        sb.append('\n');

        sb.append(String.format("Height of row 1 is: %1$s pixels.", Math.round(LengthUnitConverter.convert(ws.getRow(0).getHeight(), LengthUnit.TWIP, LengthUnit.PIXEL))));
        sb.append('\n');

        if (ws.getPictures().size() > 1) {
            ExcelPicture picture = ws.getPictures().get(1);
            sb.append(String.format("Image width x height is: %1$.2f centimeters x %2$.2f centimeters.",
                    picture.getPosition().getWidth(LengthUnit.CENTIMETER),
                    picture.getPosition().getHeight(LengthUnit.CENTIMETER)));
        }
        System.out.println(sb.toString());
    }
}