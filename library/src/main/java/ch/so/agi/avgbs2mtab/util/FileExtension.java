package ch.so.agi.avgbs2mtab.util;

import java.io.File;
import java.io.IOException;

/**
 * Utility-Class to get File extensions
 */
public class FileExtension {

    private FileExtension() {
    }

    /**
     * Gets the extension of the given file.
     *
     * @param inputFile File, which should be checked for the extension
     * @return          file extension (e.g. ".sql")
     */
    public static String getFileExtension(File inputFile)
            throws IOException {

        String [] splittedFilePath = splitFilePathAtPoint(inputFile);
        return getFileExtensionFromArray(splittedFilePath);

    }

    private static String[] splitFilePathAtPoint (File inputFile)
            throws IOException {
        String filePath =inputFile.getAbsolutePath();
        return filePath.split("\\.");
    }

    private static String getFileExtensionFromArray(String[] splittedFilePath) throws IOException{
        Integer arrayLength=splittedFilePath.length;
        if (arrayLength >= 2) {
            return splittedFilePath[arrayLength - 1];
        } else  {
            throw new IOException("File must have a file extension");
        }
    }
}



