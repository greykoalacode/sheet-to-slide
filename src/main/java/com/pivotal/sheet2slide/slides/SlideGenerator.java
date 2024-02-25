package com.pivotal.sheet2slide.slides;

import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.util.List;
import java.util.Map;

import static com.pivotal.sheet2slide.utils.FileUtils.readProperties;

public class SlideGenerator {

    private static String headerText;
    private static String headerSubText;
    private static String footerText;
    private static String footerSubText;

    public static void populateSlides(List<String[]> csvData, XMLSlideShow ppt) {
        // Internal Variables
        // Define table dimensions based on CSV data
        int numRows = csvData.size();
        int numCols = csvData.get(0).length;

        // minimum row initialization (header row + 1 row of data)
        int minRows = 2;
        // current row count
        int rowCount = 0;

        // get slide height & width
        int pageWidth = ppt.getPageSize().width;
        int pageHeight = ppt.getPageSize().height;
        // set the ratio of table : slide , i.e. 60% of slide height should be occupied by table
        int maxTableHeight = (int) (0.60 * pageHeight);

        while (rowCount < numRows - 1) {

            XSLFSlide newSlide = ppt.createSlide();

            // Create a table
            XSLFTable table = newSlide.createTable(minRows, numCols);
            String[] headerRowData = csvData.get(0);
            table.setColumnWidth(1, 60.0);
            table.setColumnWidth(numCols - 4, 200.0);
            table.setColumnWidth(numCols - 3, 70.0);
            table.setColumnWidth(numCols - 2, 40.0);
            table.setColumnWidth(numCols - 1, 40.0);

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

            // add rows till maxTableHeight is reached
            int currentRow = 1;
            while (tableHeight < 0.9 * maxTableHeight && (rowCount + currentRow) < numRows) {

                String[] rowData = csvData.get(rowCount + currentRow);
                for (int col = 0; col < numCols; col++) {
                    XSLFTableCell cell = table.getCell(currentRow, col);
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
                table.insertRow(currentRow);
            }


            // Set the position of the table to the center of the slide
            double tableWidth = 0;
            for (int col = 0; col < table.getNumberOfColumns(); col++) {
                tableWidth += table.getColumnWidth(col);
            }

            int tableXCoordinate = (pageWidth - (int) tableWidth) / 2;
            int tableYCoordinate = (pageHeight - (int) tableHeight) / 4;
            table.setAnchor(new Rectangle(tableXCoordinate, tableYCoordinate, (int) tableWidth, (int) tableHeight));

            if (loadSlideProperties()) {
                setSlideLayout(newSlide, 50 + tableYCoordinate + tableHeight);
            }

            rowCount += currentRow - 1;
        }
    }

    private static void setSlideLayout(XSLFSlide slide, double tableHeight) {
        if (headerText != null) {
            XSLFTextBox headerTextBox = slide.createTextBox();
            headerTextBox.setAnchor(new Rectangle2D.Double(50, 0, 500, 25)); // Adjust positioning and size as needed
            headerTextBox.setText(headerText);
            if (headerSubText != null) {
                XSLFTextRun subHeaderTextRun = headerTextBox.addNewTextParagraph().addNewTextRun();
                subHeaderTextRun.setFontSize(8.0);
                subHeaderTextRun.setText(headerSubText);
            }
        }

        // setting Footer
        if (footerText != null) {
            XSLFTextBox footerTextBox = slide.createTextBox();
            footerTextBox.setAnchor(new Rectangle2D.Double(50, tableHeight + 75, 500, 25)); // Adjust positioning and size as needed
            footerTextBox.setText(footerText);
            if (footerSubText != null) {
                XSLFTextRun subFooterTextRun = footerTextBox.addNewTextParagraph().addNewTextRun();
                subFooterTextRun.setFontSize(8.0);
                subFooterTextRun.setText(footerSubText);
            }
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

    private static boolean loadSlideProperties() {
        Map<Object, Object> propertiesMap = readProperties();
        if (propertiesMap != null) {
            headerText = (String) propertiesMap.get("headerText");
            headerSubText = (String) propertiesMap.get("headerSubText");
            footerText = (String) propertiesMap.get("footerText");
            footerSubText = (String) propertiesMap.get("footerSubText");
            return true;
        }
        return false;
    }

}
