package com.pivotal.sheet2slide.utils;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class FindCsvFiles {
    public static String BASE_PATH = System.getProperty("user.dir")+"/src/main/resources";
    public static List<File> readCSVFiles(String inputFolder) throws IOException {
        List<File> csvFiles = new ArrayList<>();
        File folder = new File(BASE_PATH+inputFolder);
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
