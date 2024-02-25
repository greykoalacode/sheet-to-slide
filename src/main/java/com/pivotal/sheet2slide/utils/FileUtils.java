package com.pivotal.sheet2slide.utils;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import static com.pivotal.sheet2slide.SheetToSlideApp.logger;

public class FileUtils {
    public static String propertiesFilePath = "src/main/resources/setup.properties";
    public static String BASE_PATH = System.getProperty("user.dir")+"/src/main/resources";

    public static void checkAndCreateFolder(String pathName) {
        Path parentDirectory = Paths.get(pathName).getParent();
        if (parentDirectory != null) {
            try {
                Files.createDirectories(parentDirectory);
            } catch (IOException e) {
                System.err.println("Failed to create parent directories: " + e.getMessage());
            }
        }
    }

    public static Map<Object, Object> readProperties() {
        Properties properties = new Properties();
        try (FileInputStream inputStream = new FileInputStream(propertiesFilePath)) {
            properties.load(inputStream);
            return properties;
        } catch (Exception e){
            System.out.println("ERROR: Conversion cannot take place without 'setup.properties' file.");
            logger.error(e.getMessage(), e);
            e.printStackTrace();
            return null;
        }
    }

    public static List<String[]> parseCSV(File file) throws IOException, CsvException {
//        SlideGeneratorApp.class.getResourceAsStream(fileName)
        InputStream inputStream = new FileInputStream(file);
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
        CSVReader csvReader = new CSVReader(reader);
        List<String[]> lines = csvReader.readAll();
        reader.close();
        if (lines.isEmpty()){
            return null;
        }
        return lines;
    }

    public static List<File> findCSVFiles(String inputFolder) {
        List<File> csvFiles = new ArrayList<>();
        File folder = new File(BASE_PATH+"/"+inputFolder);
        if (!folder.exists() || !folder.isDirectory()) {
            System.out.println("Folder does not exist or is a directory: " + inputFolder);
            return csvFiles;
        }
        findCSVFilesRecursive(folder, csvFiles);
        return csvFiles;
    }

    private static void findCSVFilesRecursive(File folder, List<File> csvFiles) {
        File[] filesList = folder.listFiles();
        if (filesList != null) {
            for (File file : filesList) {
                if(file.isDirectory()){
                    findCSVFilesRecursive(file, csvFiles);
                } else if(file.isFile() && file.getName().toLowerCase().endsWith(".csv")){
                    csvFiles.add(file);
                }
            }
        }
    }
}
