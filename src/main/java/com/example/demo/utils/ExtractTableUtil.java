package com.example.demo.utils;

// Spire.PDF for Java uses the PdfTableExtractor.extractTable(int pageIndex) method
// to detect and extract tables from a desired PDF page.

import com.spire.doc.*;
import com.spire.doc.documents.BorderStyle;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.pdf.PdfDocument;
import com.spire.pdf.utilities.PdfTable;
import com.spire.pdf.utilities.PdfTableExtractor;
import com.spire.xls.ExcelVersion;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.awt.*;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class ExtractTableUtil {

    public void extractTableData() throws IOException {
        // Load a sample PDF document
        PdfDocument pdf = new PdfDocument("Table.pdf");

        // Create a StringBuilder instance
        StringBuilder builder = new StringBuilder();

        // Create a PdfTableExtractor instance
        PdfTableExtractor extractor = new PdfTableExtractor(pdf);

        // Loop through the pages in the PDF
        for (int pageIndex = 0; pageIndex < pdf.getPages().getCount(); pageIndex++) {
            // Extract tables from the current page into a PdfTable array
            PdfTable[] tableLists = extractor.extractTable(pageIndex);

            // If any tables are found
            if (tableLists != null && tableLists.length > 0) {
                // Loop through the tables in the array
                for (PdfTable table : tableLists) {
                    // Loop through the rows in the current table
                    for (int i = 0; i < table.getRowCount(); i++) {
                        // Loop through the columns in the current table
                        for (int j = 0; j < table.getColumnCount(); j++) {
                            // Extract data from the current table cell and append to the StringBuilder
                            String text = table.getText(i, j);
                            builder.append(text + " | ");
                        }
                        builder.append("\r\n");
                    }
                }
            }
        }

        // Write data into a .txt document
        FileWriter fw = new FileWriter("ExtractTable.txt");
        fw.write(builder.toString());
        fw.flush();
        fw.close();
    }

    public void extractTableDataToWord() throws IOException {
        // Load a sample PDF document containing tables
        Resource resource = new ClassPathResource("table.pdf");
        File file = resource.getFile();
        PdfDocument pdf = new PdfDocument(file.getAbsolutePath());

        // Create a PdfTableExtractor instance
        PdfTableExtractor extractor = new PdfTableExtractor(pdf);

        // Extract tables from a specific page
        PdfTable[] pdfTables  = extractor.extractTable(0);

        // Create a Document object
        Document doc = new Document();

        // Add a section
        Section section = doc.addSection();

        // If any tables are found
        if (pdfTables != null && pdfTables.length > 0) {

            // Loop through the tables
            for (int tableNum = 0; tableNum < pdfTables.length; tableNum++) {

                //Get a specific PDF table
                PdfTable  pdfTable = pdfTables[tableNum];

                //Add a table in Word
                Table wordTable = section.addTable();
                wordTable.resetCells(pdfTable.getRowCount(), pdfTable.getColumnCount());
                wordTable.setPreferredWidth(PreferredWidth.getAuto());
                wordTable.getTableFormat().getBorders().setBorderType(BorderStyle.Single);
                wordTable.getTableFormat().getBorders().setColor(Color.BLACK);

                //Loop through the rows in the current PDF table
                for (int rowNum = 0; rowNum < pdfTable.getRowCount(); rowNum++) {

                    //Loop through the columns in the current PDF table
                    for (int colNum = 0; colNum < pdfTable.getColumnCount(); colNum++) {

                        //Extract data from the current table cell
                        String text = pdfTable.getText(rowNum, colNum);

                        //Write data to Word table
                        Paragraph paragraph = wordTable.get(rowNum, colNum).addParagraph();
                        paragraph.appendText(text);
                        paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                    }
                }

                //Add a blank paragraph
                section.addParagraph();
            }
        }

        //Save the document to a docx file
        doc.saveToFile("output/ToWord.docx", FileFormat.Docx);
    }

    public void extractTableDataToExcel() {
        //Load a sample PDF document containing tables
        PdfDocument pdf = new PdfDocument("C:\\Users\\Administrator\\Desktop\\Tables.pdf");

        //Create a PdfTableExtractor instance
        PdfTableExtractor extractor = new PdfTableExtractor(pdf);

        //Extract tables from a specific page
        PdfTable[] pdfTables  = extractor.extractTable(0);

        //Create a Workbook object,
        Workbook wb = new Workbook();

        //Remove default worksheets
        wb.getWorksheets().clear();

        //If any tables are found
        if (pdfTables != null && pdfTables.length > 0) {

            //Loop through the tables
            for (int tableNum = 0; tableNum < pdfTables.length; tableNum++) {

                //Add a worksheet to workbook
                String sheetName = String.format("Table - %d", tableNum + 1);
                Worksheet sheet = wb.getWorksheets().add(sheetName);

                //Loop through the rows in the current table
                for (int rowNum = 0; rowNum < pdfTables[tableNum].getRowCount(); rowNum++) {

                    //Loop through the columns in the current table
                    for (int colNum = 0; colNum < pdfTables[tableNum].getColumnCount(); colNum++) {

                        //Extract data from the current table cell
                        String text = pdfTables[tableNum].getText(rowNum, colNum);

                        //Insert data into a specific cell
                        sheet.get(rowNum + 1, colNum + 1).setText(text);

                    }
                }

                //Autofit column width
                for (int sheetColNum = 0; sheetColNum < sheet.getColumns().length; sheetColNum++) {
                    sheet.autoFitColumn(sheetColNum + 1);
                }
            }
        }

        //Save the workbook to an Excel file
        wb.saveToFile("output/ToExcel.xlsx", ExcelVersion.Version2016);

    }
}
