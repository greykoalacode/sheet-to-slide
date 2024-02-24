package com.pivotal.sheet2slide;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Hello world!
 */
public class App {
    private static final Logger logger = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) throws Exception {
        // Load CSV data
        List<String[]> csvData = readCSV("/MOCK_DATA.csv");
        long startTime = System.nanoTime();

        try (OutputStream os = Files.newOutputStream(Paths.get("src/main/resources/sample.pptx")); XMLSlideShow ppt = new XMLSlideShow()) {


            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

            XSLFSlide slide = ppt.createSlide(layout);

            XSLFTextShape titleShape = slide.getPlaceholder(0);
            XSLFTextShape contentShape = slide.getPlaceholder(1);
            titleShape.setText("Title of Slides");
            contentShape.setText("Content is here");

            // Define table dimensions based on CSV data
            int numRows = csvData.size();
            int numCols = csvData.get(0).length;

            int maxRowsPerSlide = 3;

            int minRows = 2;
            int rowCount = 0;

            int pageWidth = ppt.getPageSize().width;
            int pageHeight = ppt.getPageSize().height;

            int maxTableHeight = (int) (0.75 * pageHeight);

            while (rowCount < numRows-1) {

                XSLFSlide newSlide = ppt.createSlide();
                // Create a table
                XSLFTable table = newSlide.createTable(minRows, numCols);
                String[] headerRowData = csvData.get(0);
                table.setColumnWidth(2, 200.0);
                table.setColumnWidth(5, 40.0);

                double tableHeight = 0;

                // header row
                for (int headerCol = 0; headerCol < numCols; headerCol++) {
                    XSLFTableCell cell = table.getCell(0, headerCol);
                    setCellBorder(cell);
                    XSLFTextParagraph p = cell.addNewTextParagraph();
                    p.setTextAlign(TextParagraph.TextAlign.CENTER);
                    XSLFTextRun r = p.addNewTextRun();
                    r.setFontColor(Color.WHITE);
                    r.setFontFamily("Arial");
                    r.setFontSize(8.0);
                    r.setBold(true);
                    cell.setFillColor(new Color(0, 0, 128));
                    r.setText(headerRowData[headerCol]);
                    tableHeight += table.getRowHeight(0);
                }

//                add rows till maxTableHeight is reached
                int currentRow = 1;
                while (tableHeight < maxTableHeight && (rowCount+currentRow) < numRows) {

                    String[] rowData = csvData.get(rowCount + currentRow);
                    for (int col = 0; col < numCols; col++) {
//
//                        row.addCell();
                        XSLFTableCell cell = table.getCell(currentRow, col);;
                        System.out.println("rowCount "+rowCount+" currentRow "+currentRow+" col "+col);
                        setCellBorder(cell);
                        XSLFTextParagraph p = cell.addNewTextParagraph();
                        XSLFTextRun r = p.addNewTextRun();
                        r.setText(rowData[col]);
                        r.setFontFamily("Arial");
                        r.setFontSize(8.0);
                        r.setText(rowData[col]);
                    }
                    tableHeight += table.getRowHeight(currentRow);
                    currentRow += 1;
//                    XSLFTableRow row = table.addRow();
                    table.insertRow(currentRow);
                }

//                int availableRows = Math.min(maxRowsPerSlide + 1, csvData.size() - rowCount);
                // Populate the table with CSV data
//                for (int currentRow = 1; currentRow < availableRows; currentRow++) {
//                    String[] rowData = csvData.get(rowCount + currentRow);
//                    for (int col = 0; col < numCols; col++) {
//                        XSLFTableCell cell = table.getCell(currentRow, col);
//                        setCellBorder(cell);
//                        XSLFTextParagraph p = cell.addNewTextParagraph();
//                        XSLFTextRun r = p.addNewTextRun();
//                        r.setText(rowData[col]);
//                        r.setFontFamily("Arial");
//                        r.setFontSize(8.0);
//                        r.setText(rowData[col]);
//
//
//                    }
//                    tableHeight += table.getRowHeight(currentRow);
//                }


                // Set the position of the table to the center of the slide
                double tableWidth = 0;
                for (int col = 0; col < table.getNumberOfColumns(); col++) {
                    tableWidth += table.getColumnWidth(col);
                }
                int tableXCoordinate = (pageWidth - (int) tableWidth) / 2;
                int tableYCoordinate = (pageHeight - (int) tableHeight) / 2;
                table.setAnchor(new Rectangle(tableXCoordinate, tableYCoordinate, (int) tableWidth, (int) tableHeight));


                rowCount += currentRow-1;
            }
            ppt.write(os);
            long endTime = System.nanoTime();
            long totalTime = endTime - startTime;
            System.out.printf("Script runtime: %s s%n", convertNanoToSeconds(totalTime));
        } catch (Exception e) {
            logger.error(e.getMessage());
            e.printStackTrace();
        }
    }

    private static void setCellBorder(XSLFTableCell cell) {
        XSLFTableCell.BorderEdge[] borderEdges = {
                XSLFTableCell.BorderEdge.bottom,
                XSLFTableCell.BorderEdge.left,
                XSLFTableCell.BorderEdge.right,
                XSLFTableCell.BorderEdge.top
        };
        for (XSLFTableCell.BorderEdge borderEdge : borderEdges) {
            cell.setBorderColor(borderEdge, Color.BLACK);
        }
    }

    private static List<String[]> readCSV(String fileName) throws IOException, CsvException {
        InputStream inputStream = App.class.getResourceAsStream(fileName);
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
        CSVReader csvReader = new CSVReader(reader);
        List<String[]> lines = csvReader.readAll();
        reader.close();
        return lines;
    }

    private static double convertNanoToSeconds(long durationInNano) {
        return TimeUnit.MILLISECONDS.convert(durationInNano, TimeUnit.NANOSECONDS) * 0.001;
    }
}
