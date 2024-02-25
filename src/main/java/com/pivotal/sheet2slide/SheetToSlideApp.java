package com.pivotal.sheet2slide;

import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import static com.pivotal.sheet2slide.slides.SlideGenerator.populateSlides;
import static com.pivotal.sheet2slide.utils.FileUtils.*;

/**
 * Slide Generator Class
 */
public class SheetToSlideApp {
    public static final Logger logger = LoggerFactory.getLogger(SheetToSlideApp.class);

    private static String inputFolder;
    private static String outputFolder;


    public static void main(String[] args) throws Exception {
        long startTime = System.nanoTime();

        // read properties file
        if (initializeSetup()) {
            // Load CSV data
            List<File> csvFiles = findCSVFiles(inputFolder);

            if (!csvFiles.isEmpty()) {
                // Convert CSV -> Slides
                for (File eachCsvFile : csvFiles) {
                    List<String[]> csvData = parseCSV(eachCsvFile);
                    if(csvData != null) {
                        String outputFileName = eachCsvFile.getName().replace(".csv", ".pptx");

                        System.out.println("Creating PowerPoint from CSV...");
                        System.out.println("Input CSV file: " + eachCsvFile.getName());
                        System.out.println("Output PPTX file: " + outputFileName);

                        String outputFilePath = BASE_PATH + "/"+outputFolder + "/" + outputFileName;
                        checkAndCreateFolder(outputFilePath);
                        try (OutputStream os = Files.newOutputStream(Paths.get(outputFilePath)); XMLSlideShow ppt = new XMLSlideShow()) {
                            // PPT Initiation
                            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
                            XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);

                            // Creating First Slide with the layout - Title & Subtitle
                            XSLFSlide slide = ppt.createSlide(layout);
                            XSLFTextShape titleShape = slide.getPlaceholder(0);
                            XSLFTextShape contentShape = slide.getPlaceholder(1);
                            titleShape.setText("Title of Slides");
                            contentShape.setText("Content is here");

                            // Populate Slides based on data
                            populateSlides(csvData, ppt);

                            ppt.write(os);

                        } catch (Exception e) {
                            logger.error(e.getMessage());
                            e.printStackTrace();
                        }
                    } else {
                        logger.error(String.format("There is no data in CSV File %s. Check 'resources/input' folder.", eachCsvFile.getName()));
                    }
                }
            } else {
                logger.error("There are no CSV File(s). Check 'resources/input' folder.");
            }

        } else {
            logger.error("'setup.properties' file does not have any properties. Please add according to the manual.");
        }

        long endTime = System.nanoTime();
        long totalTime = endTime - startTime;
        System.out.printf("Script runtime: %.3f s%n", convertNanoToSeconds(totalTime));
    }

    private static boolean initializeSetup() {
        Map<Object, Object> propertiesMap = readProperties();
        if (propertiesMap != null) {
            inputFolder = (String) propertiesMap.get("inputFolder");
            outputFolder = (String) propertiesMap.get("outputFolder");
            return inputFolder != null && outputFolder != null;
        }
        return false;
    }


    private static void printUsage() {
        System.out.println("Usage: java SheetToSlideApp -i <input CSV file> -o <output PPTX file>");
        System.out.println("Options:");
        System.out.println("  -i, --input    Specify the input CSV file path.");
        System.out.println("  -o, --output   Specify the output PPTX file path.");
        System.out.println("  -h, --help     Display this help message.");
    }


    private static double convertNanoToSeconds(long durationInNano) {
        return TimeUnit.MILLISECONDS.convert(durationInNano, TimeUnit.NANOSECONDS) * 0.001;
    }
}
