package com.pivotal.sheet2slide;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static com.pivotal.sheet2slide.utils.FindCsvFiles.BASE_PATH;
import static com.pivotal.sheet2slide.utils.FindCsvFiles.readCSVFiles;

/**
 * Hello world!
 */
public class SlideGeneratorApp {
    private static final Logger logger = LoggerFactory.getLogger(SlideGeneratorApp.class);

    private static final String inputFolder = "/input";
    private static final String outputFolder = "/output";

    public static void main(String[] args) throws Exception {
        // input

        long startTime = System.nanoTime();
        // Load CSV data
        List<File> csvFiles = readCSVFiles(inputFolder);
        for(File eachCsvFile: csvFiles){
            List<String[]> csvData = readCSV(eachCsvFile);
            String outputFilePath = BASE_PATH+outputFolder+"/"+eachCsvFile.getName().replace(".csv",".pptx");
            checkAndCreateFolder(outputFilePath);
            try (OutputStream os = Files.newOutputStream(Paths.get(outputFilePath)); XMLSlideShow ppt = new XMLSlideShow()) {
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

                int minRows = 2;
                int rowCount = 0;

                int pageWidth = ppt.getPageSize().width;
                int pageHeight = ppt.getPageSize().height;

                int maxTableHeight = (int) (0.60 * pageHeight);

                while (rowCount < numRows - 1) {

                    XSLFSlide newSlide = ppt.createSlide();
                    setSlideLayout(newSlide);
                    // Create a table
                    XSLFTable table = newSlide.createTable(minRows, numCols);
                    String[] headerRowData = csvData.get(0);
                    table.setColumnWidth(1, 60.0);
                    table.setColumnWidth(numCols-4, 200.0);
                    table.setColumnWidth(numCols-3, 70.0);
                    table.setColumnWidth(numCols-2, 40.0);
                    table.setColumnWidth(numCols-1, 40.0);

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
                    while (tableHeight < 0.9*maxTableHeight && (rowCount + currentRow) < numRows) {

                        String[] rowData = csvData.get(rowCount + currentRow);
                        for (int col = 0; col < numCols; col++) {
                            XSLFTableCell cell = table.getCell(currentRow, col);
                            System.out.println("rowCount " + rowCount + " currentRow " + currentRow + " col " + col);
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
//                    double actualTableHeight = 0;
                    for (int col = 0; col < table.getNumberOfColumns(); col++) {
                        tableWidth += table.getColumnWidth(col);
                    }
                    int tableXCoordinate = (pageWidth - (int) tableWidth) / 2;
                    int tableYCoordinate = (pageHeight - (int) tableHeight) / 4;
                    table.setAnchor(new Rectangle(tableXCoordinate, tableYCoordinate, (int) tableWidth, (int) tableHeight));


                    rowCount += currentRow - 1;
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

    }

    private static void setSlideLayout(XSLFSlide slide){
        String headerText = "This is Header";
        String headerSubText = "Issues listed from 16-02-2024 to 24-02-2024";
//        XSLFAutoShape headerShape = slide.createAutoShape();
//        headerShape.setText("This is Header");
//        headerShape.addNewTextParagraph().addNewTextRun().setText("issues listed from 16-02-2024 to 24-02-2024");

        XSLFTextBox headerTextBox = slide.createTextBox();
        headerTextBox.setAnchor(new Rectangle2D.Double(50, 0, 500, 25)); // Adjust positioning and size as needed
        headerTextBox.setText(headerText);
        XSLFTextRun subHeaderTextRun = headerTextBox.addNewTextParagraph().addNewTextRun();
        subHeaderTextRun.setFontSize(8.0);
        subHeaderTextRun.setText(headerSubText);
    }

    private static void checkAndCreateFolder(String pathName){
        Path parentDirectory = Paths.get(pathName).getParent();
        if (parentDirectory != null) {
            try {
                Files.createDirectories(parentDirectory);
            } catch (IOException e) {
                System.err.println("Failed to create parent directories: " + e.getMessage());
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

    private static void printUsage() {
        System.out.println("Usage: java PowerPointCLI -i <input CSV file> -o <output PPTX file>");
        System.out.println("Options:");
        System.out.println("  -i, --input    Specify the input CSV file path.");
        System.out.println("  -o, --output   Specify the output PPTX file path.");
        System.out.println("  -h, --help     Display this help message.");
    }

    private static void createPowerPointFromCSV(String inputCsvFile, String outputPptxFile) throws Exception {
        // Your existing code to create PowerPoint from CSV goes here
        // Remember to replace placeholders with actual implementation
        System.out.println("Creating PowerPoint from CSV...");
        System.out.println("Input CSV file: " + inputCsvFile);
        System.out.println("Output PPTX file: " + outputPptxFile);
        // Call your existing code here
    }

    private static List<String[]> readCSV(File file) throws IOException, CsvException {
//        SlideGeneratorApp.class.getResourceAsStream(fileName)
        InputStream inputStream = new FileInputStream(file);
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
