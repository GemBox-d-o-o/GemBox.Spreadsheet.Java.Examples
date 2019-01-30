package com.gembox.examples.spring.controllers;

import com.gembox.spreadsheet.*;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Controller
@RequestMapping("workbook")
public class WorkbookController {

    static {
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");
    }

    private static final List<WorkbookItem> DATA = Arrays.asList(
        new WorkbookItem(100, "John", "Doe"),
        new WorkbookItem(101, "Fred", "Nurk"),
        new WorkbookItem(102, "Hans", "Meier"),
        new WorkbookItem(103, "Ivan", "Horvat"),
        new WorkbookItem(104, "Jean", "Dupont"),
        new WorkbookItem(105, "Mario", "Rossi")
    );

    @RequestMapping(value = "/create", method = RequestMethod.GET)
    public String create(Model model) {
        model.addAttribute("workbookItemsWithFormat", new WorkbookItemsWithFormat("XLSX", DATA));
        return "create";
    }

    @RequestMapping(value = "/create", method = RequestMethod.POST)
    public HttpEntity<byte[]> create(@ModelAttribute("workbookItemsWithFormat") WorkbookItemsWithFormat workbookItemsWithFormat) throws IOException {

        SaveOptions options = getSaveOptions(workbookItemsWithFormat.getSelectedFormat());
        ExcelFile book = new ExcelFile();
        ExcelWorksheet sheet = book.addWorksheet("Sheet1");

        CellStyle style = sheet.getRow(0).getStyle();
        style.getFont().setWeight(ExcelFont.BOLD_WEIGHT);
        style.setHorizontalAlignment(HorizontalAlignmentStyle.CENTER);
        sheet.getColumn(0).getStyle().setHorizontalAlignment(HorizontalAlignmentStyle.CENTER);

        sheet.getColumn(0).setWidth(50, LengthUnit.PIXEL);
        sheet.getColumn(1).setWidth(150, LengthUnit.PIXEL);
        sheet.getColumn(2).setWidth(150, LengthUnit.PIXEL);

        sheet.getCell("A1").setValue("ID");
        sheet.getCell("B1").setValue("First Name");
        sheet.getCell("C1").setValue("Last Name");

        for (int row = 1; row <= workbookItemsWithFormat.getItems().size(); row++) {
            WorkbookItem item = workbookItemsWithFormat.getItems().get(row - 1);
            sheet.getCell(row, 0).setValue(item.getId());
            sheet.getCell(row, 1).setValue(item.getFirstName());
            sheet.getCell(row, 2).setValue(item.getLastName());
        }

        byte[] bytes = getBytes(book, options);

        HttpHeaders header = new HttpHeaders();
        header.set(HttpHeaders.CONTENT_TYPE, options.getContentType());
        header.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Create." + workbookItemsWithFormat.getSelectedFormat().toLowerCase());
        header.setContentLength(bytes.length);

        return new HttpEntity<>(bytes, header);
    }

    private byte[] getBytes(ExcelFile book, SaveOptions options) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        book.save(outputStream, options);
        return outputStream.toByteArray();
    }

    private static SaveOptions getSaveOptions(String format) {
        switch (format.toUpperCase()) {
            case "XLSX":
                return SaveOptions.getXlsxDefault();
            case "XLS":
                return SaveOptions.getXlsDefault();
            case "ODS":
                return SaveOptions.getOdsDefault();
            case "CSV":
                return SaveOptions.getCsvDefault();
            case "HTML":
                return SaveOptions.getHtmlDefault();
            default:
                throw new IllegalArgumentException("Format '" + format + "' is not supported.");
        }
    }


    public static class WorkbookItem {

        private int id;
        private String firstName;
        private String lastName;

        public WorkbookItem(int id, String firstName, String lastName) {
            this.id = id;
            this.firstName = firstName;
            this.lastName = lastName;
        }

        public WorkbookItem() {
        }

        public int getId() {
            return id;
        }

        public void setId(int id) {
            this.id = id;
        }

        public String getFirstName() {
            return firstName;
        }

        public void setFirstName(String firstName) {
            this.firstName = firstName;
        }

        public String getLastName() {
            return lastName;
        }

        public void setLastName(String lastName) {
            this.lastName = lastName;
        }
    }

    public static class WorkbookItemsWithFormat {

        public String selectedFormat;
        public List<WorkbookItem> items = new ArrayList<>();

        public WorkbookItemsWithFormat(String selectedFormat, List<WorkbookItem> items) {
            this.selectedFormat = selectedFormat;
            this.items = items;
        }

        public WorkbookItemsWithFormat() {
        }

        public String getSelectedFormat() {
            return selectedFormat;
        }

        public void setSelectedFormat(String selectedFormat) {
            this.selectedFormat = selectedFormat;
        }

        public List<WorkbookItem> getItems() {
            return items;
        }

        public void setItems(List<WorkbookItem> items) {
            this.items = items;
        }
    }
}
