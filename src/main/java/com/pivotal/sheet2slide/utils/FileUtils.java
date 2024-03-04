package com.pivotal.sheet2slide.utils;

import com.pivotal.sheet2slide.exceptions.CsvFilesNotFoundException;
import com.pivotal.sheet2slide.exceptions.FileEmptyException;
import com.pivotal.sheet2slide.exceptions.FolderNotExistException;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.logging.Logger;
import java.util.stream.Collectors;


public class FileUtils {
    public static String propertiesFilePath = "src/main/resources/setup.properties";
    public static String BASE_PATH = System.getProperty("user.dir") + "/src/main/resources";

    public static final Logger logger = Logger.getLogger("FileUtils-Logger");


    public static boolean isRequiredPropertiesEmpty(Map<Object, Object> properties) {
        return properties == null || isNullOrEmpty((String) properties.get("inputFolder")) || isNullOrEmpty((String) properties.get("outputFolder"));
    }

    public static boolean isNullOrEmpty(String value) {
        return value == null || value.isEmpty();
    }

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
        } catch (Exception e) {
            logger.severe("Could not find the mentioned file: " + propertiesFilePath);
        }
        return null;
    }

    public static List<String[]> parseCSV(File file) throws IOException, FileEmptyException {
        InputStream inputStream = new FileInputStream(file);
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
        CSVParser parser = new CSVParser(reader, CSVFormat.EXCEL);
        List<CSVRecord> csvRecords = parser.getRecords();
        List<String[]> lines = new ArrayList<>();
        for(CSVRecord record: csvRecords){
            lines.add(record.values());
        }
        reader.close();
        if (lines.isEmpty()) {
            throw new FileEmptyException(String.format("CSV File %s is Empty", file.getName()));
        }
        return lines;
    }

    public static List<File> findCSVFiles(String inputFolder) throws FolderNotExistException, CsvFilesNotFoundException {
        List<File> csvFiles = new ArrayList<>();
        File folder = new File(BASE_PATH + "/" + inputFolder);
        if (!folder.exists() || !folder.isDirectory()) {
            throw new FolderNotExistException("Folder does not exist or is a directory: " + inputFolder);
        }
        findCSVFilesRecursive(folder, csvFiles);
        if (csvFiles.isEmpty()) {
            logger.severe("There are no CSV File(s). Check 'resources/input' folder.");
            throw new CsvFilesNotFoundException("No CSV Files found in folder: " + inputFolder);
        }
        return csvFiles;
    }

    private static void findCSVFilesRecursive(File folder, List<File> csvFiles) {
        File[] filesList = folder.listFiles();
        if (filesList != null) {
            for (File file : filesList) {
                if (file.isDirectory()) {
                    findCSVFilesRecursive(file, csvFiles);
                } else if (file.isFile() && file.getName().toLowerCase().endsWith(".csv")) {
                    csvFiles.add(file);
                }
            }
        }
    }
}
