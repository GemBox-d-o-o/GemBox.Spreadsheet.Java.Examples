import com.gembox.spreadsheet.*;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Images");

        ws.getCell(0, 0).setValue("Image examples:");

        // Small BMP added by using rectangle.
        ws.getPictures().add("SmallImage.bmp", 50, 50, 48, 48, LengthUnit.PIXEL);

        // Large JPG added by using one anchor.
        ws.getPictures().add("FragonardReader.jpg", "B9");

        // PNG added by using two anchors.
        ws.getPictures().add("Dices.png", "J16", "K20");

        // GIF added by using anchors. Notice that animation is lost in MS Excel.
        ws.getPictures().add("Zahnrad.gif",
                new AnchorCell(ws.getColumn(9), ws.getRow(21), 100000, 100000),
                new AnchorCell(ws.getColumn(10), ws.getRow(23), 50000, 50000)).getPosition().setMode(PositioningMode.MOVE);

        // WMF added by using one anchor and size.
        ws.getPictures().add("Graphics1.wmf", "J9", 250, 100, LengthUnit.PIXEL).getPosition().setMode(PositioningMode.FREE_FLOATING);

        ef.save("Images.xlsx");
    }
}