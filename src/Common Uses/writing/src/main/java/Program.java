import com.gembox.spreadsheet.*;

import java.awt.*;
import java.util.EnumSet;

class Program {

    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.addWorksheet("Writing");

        // Tabular sample data for writing into an Excel file.
        Object[][] skyscrapers = {
                {"Rank", "Building", "City", "Metric", "Imperial", "Floors", "Built (Year)"},
                { 1, "Taipei 101", "Taipei", 509, 1671, 101, 2004},
                { 2, "Petronas Tower 1", "Kuala Lumpur", 452, 1483, 88, 1998},
                { 3, "Petronas Tower 2", "Kuala Lumpur", 452, 1483, 88, 1998},
                { 4, "Sears Tower", "Chicago", 442, 1450, 108, 1974},
                { 5, "Jin Mao Tower", "Shanghai", 421, 1380, 88, 1998},
                { 6, "2 International Finance Centre", "Hong Kong", 415, 1362, 88, 2003},
                { 7, "CITIC Plaza", "Guangzhou", 391, 1283, 80, 1997},
                { 8, "Shun Hing Square", "Shenzhen", 384, 1260, 69, 1996},
                { 9, "Empire State Building", "New York City", 381, 1250, 102, 1931},
                {10, "Central Plaza", "Hong Kong", 374, 1227, 78, 1992},
                {11, "Bank of China Tower", "Hong Kong", 367, 1205, 72, 1990},
                {12, "Emirates Office Tower", "Dubai", 355, 1163, 54, 2000},
                {13, "Tuntex Sky Tower", "Kaohsiung", 348, 1140, 85, 1997},
                {14, "Aon Center", "Chicago", 346, 1136, 83, 1973},
                {15, "The Center", "Hong Kong", 346, 1135, 73, 1998},
                {16, "John Hancock Center", "Chicago", 344, 1127, 100, 1969},
                {17, "Ryugyong Hotel", "Pyongyang", 330, 1083, 105, 1992},
                {18, "Burj Al Arab", "Dubai", 321, 1053, 60, 1999},
                {19, "Chrysler Building", "New York City", 319, 1046, 77, 1930},
                {20, "Bank of America Plaza", "Atlanta", 312, 1023, 55, 1992}
        };

        ws.getCell(0, 0).setValue("Example of writing typical table - tallest buildings in the world (2004):");

        // Column width of 8, 30, 16, 9, 9, 9, 9, 4 and 5 characters.
        ws.getColumn(0).setWidth(8 * 256);
        ws.getColumn(1).setWidth(30 * 256);
        ws.getColumn(2).setWidth(16 * 256);
        ws.getColumn(3).setWidth(9 * 256);
        ws.getColumn(4).setWidth(9 * 256);
        ws.getColumn(5).setWidth(9 * 256);
        ws.getColumn(6).setWidth(9 * 256);
        ws.getColumn(7).setWidth(4 * 256);
        ws.getColumn(8).setWidth(5 * 256);

        int i, j;
        // Write header data to Excel cells.
        for (j = 0; j < 7; j++)
            ws.getCell(3, j).setValue(skyscrapers[0][j]);

        ws.getCells().getSubrangeAbsolute(2, 0, 3, 0).setMerged(true);
        ws.getCells().getSubrangeAbsolute(2, 1, 3, 1).setMerged(true);
        ws.getCells().getSubrangeAbsolute(2, 2, 3, 2).setMerged(true);
        ws.getCells().getSubrangeAbsolute(2, 3, 2, 4).setMerged(true);
        ws.getCell(2, 3).setValue("Height");
        ws.getCells().getSubrangeAbsolute(2, 5, 3, 5).setMerged(true);
        ws.getCells().getSubrangeAbsolute(2, 6, 3, 6).setMerged(true);

        CellStyle tmpStyle = new CellStyle();
        tmpStyle.setHorizontalAlignment(HorizontalAlignmentStyle.CENTER);
        tmpStyle.setVerticalAlignment(VerticalAlignmentStyle.CENTER);
        tmpStyle.getFillPattern().setSolid(SpreadsheetColor.fromColor(Color.ORANGE));
        tmpStyle.getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        tmpStyle.getFont().setColor(SpreadsheetColor.fromColor(Color.WHITE));
        tmpStyle.setWrapText(true);
        tmpStyle.getBorders().setBorders(EnumSet.of(MultipleBorders.RIGHT, MultipleBorders.TOP), SpreadsheetColor.fromColor(Color.BLACK), LineStyle.THIN);

        ws.getCells().getSubrangeAbsolute(2, 0, 3, 6).setStyle(tmpStyle);

        tmpStyle = new CellStyle();
        tmpStyle.setHorizontalAlignment(HorizontalAlignmentStyle.CENTER);
        tmpStyle.setVerticalAlignment(VerticalAlignmentStyle.CENTER);
        tmpStyle.getFont().setWeight(ExcelFont.BOLD_WEIGHT);

        CellRange mergedRange = ws.getCells().getSubrangeAbsolute(4, 7, 13, 7);
        mergedRange.setMerged(true);
        mergedRange.setValue("T o p   1 0");
        tmpStyle.setRotation(-90);
        tmpStyle.getFillPattern().setSolid(SpreadsheetColor.fromName(ColorName.LIGHT_GREEN));
        mergedRange.setStyle(tmpStyle);

        mergedRange = ws.getCells().getSubrangeAbsolute(4, 8, 23, 8);
        mergedRange.setMerged(true);
        mergedRange.setValue("T o p   2 0");
        tmpStyle.setTextVertical(true);
        tmpStyle.getFillPattern().setSolid(SpreadsheetColor.fromName(ColorName.YELLOW));
        mergedRange.setStyle(tmpStyle);

        mergedRange = ws.getCells().getSubrangeAbsolute(14, 7, 23, 7);
        mergedRange.setMerged(true);
        mergedRange.setStyle(tmpStyle);

        // Write and format sample data to Excel cells.
        for (i = 0; i < 20; i++)
            for (j = 0; j < 7; j++)
            {
                ExcelCell cell = ws.getCell(i + 4, j);

                cell.setValue(skyscrapers[i + 1][j]);

                if (i % 2 == 0)
                    cell.getStyle().getFillPattern().setSolid(SpreadsheetColor.fromName(ColorName.LIGHT_BLUE));
                else
                    cell.getStyle().getFillPattern().setSolid(SpreadsheetColor.fromArgb(210, 210, 230));

                if (j == 3)
                    cell.getStyle().setNumberFormat("#\" m\"");

                if (j == 4)
                    cell.getStyle().setNumberFormat("#\" ft\"");

                if (j > 2)
                    cell.getStyle().getFont().setName("Courier New");

                cell.getStyle().getBorders().get(IndividualBorder.RIGHT).setLineStyle(LineStyle.THIN);
            }

        ws.getCells().getSubrange("A5", "I24").getStyle().getBorders().setBorders(MultipleBorders.outside(), SpreadsheetColor.fromColor(Color.BLACK), LineStyle.DOUBLE);
        ws.getCells().getSubrange("A3", "G4").getStyle().getBorders().setBorders(EnumSet.of(MultipleBorders.LEFT, MultipleBorders.RIGHT, MultipleBorders.TOP), SpreadsheetColor.fromColor(Color.BLACK), LineStyle.DOUBLE);
        ws.getCells().getSubrange("A5", "H14").getStyle().getBorders().setBorders(EnumSet.of(MultipleBorders.BOTTOM, MultipleBorders.RIGHT), SpreadsheetColor.fromColor(Color.BLACK), LineStyle.DOUBLE);

        ws.getCell("A27").setValue("Notes:");
        ws.getCell("A28").setValue("a) \"Metric\" and \"Imperial\" columns use custom number formatting.");
        ws.getCell("A29").setValue("b) All number columns use \"Courier New\" font for improved number readability.");
        ws.getCell("A30").setValue("c) Multiple merged ranges were used for table header and categories header.");

        ws.getPrintOptions().setFitWorksheetWidthToPages(1);

        ef.save("Writing.xlsx");
    }
}