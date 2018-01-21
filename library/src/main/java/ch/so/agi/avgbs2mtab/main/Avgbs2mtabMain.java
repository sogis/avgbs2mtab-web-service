package ch.so.agi.avgbs2mtab.main;

import ch.so.agi.avgbs2mtab.mutdat.*;
import ch.so.agi.avgbs2mtab.readxtf.ReadXtf;

import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;

import java.io.File;
import ch.so.agi.avgbs2mtab.writeexcel.XlsxWriter;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.IOException;
import java.util.logging.Logger;

/**
 * Main class
 */
public class Avgbs2mtabMain {

    private static Logger log = Logger.getLogger(Avgbs2mtabMain.class.getName());

    public static void main(String[] args){
        CommandlineParser cp = new CommandlineParser(args);

        //set loglevel

        //run Conversion with input- and outputfile

        //Output of Errors to the Commandline
    }

    public static void runConversion(String inputFilePath, String outputFilePath) throws IOException {
        assertValidInputFile(inputFilePath);
        assertValidOutputFilePath(outputFilePath);

        ParcelContainer parceldump = new ParcelContainer();
        DPRContainer dprdump = new DPRContainer();
        XlsxWriter xlsxWriter = new XlsxWriter(parceldump, dprdump, parceldump, dprdump);

        ReadXtf xtfreader = new ReadXtf((SetParcel)parceldump, (SetDPR)dprdump);
        xtfreader.readFile(inputFilePath);
        xlsxWriter.writeXlsx(outputFilePath);
    }

    private static void assertValidOutputFilePath(String outputFilePath){
        if(outputFilePath == null)
            throw new IllegalArgumentException("outputFilePath must not be null");

        File outputFile = new File(outputFilePath);
        if(outputFile.exists())
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_FILE_EXISTS, "outputFile can't be created because it exists already: " + outputFilePath);

        if(!outputFile.getParentFile().canWrite())
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_FOLDER_NOT_WRITEABLE, "Can't write to folder: " + outputFile.getParentFile().getAbsolutePath());
    }

    private static void assertValidInputFile(String inputFilePath){
        if(inputFilePath == null)
            throw new IllegalArgumentException("inputFilePath must not be null");

        File inputFile = new File(inputFilePath);
        if(!inputFile.isFile())
            throw new IllegalArgumentException("inputFilePath is not a file: " + inputFilePath);

        if(!inputFile.isAbsolute())
            throw new IllegalArgumentException("inputFilePath is not absolute: " + inputFilePath);

        if(!inputFile.canRead())
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_FILE_NOT_READABLE, "Can not read file at path: " + inputFilePath);

        String fileName = inputFile.getName();
        int lastPointIdx = fileName.lastIndexOf(".");
        if(lastPointIdx < 1)
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_WRONG_EXTENSION, "input file name has no extension: " + fileName);

        String[] parts = fileName.split("\\."); // . is regex special character -> needs escaping with \\
        String extension = parts[parts.length - 1];
        boolean isXTF = extension.equalsIgnoreCase("xtf");
        boolean isXML = extension.equalsIgnoreCase("xml");

        if(!isXTF && !isXML)
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_WRONG_EXTENSION, "input file extension must be xml or xtf: " + fileName);

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

        try {
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(inputFile);
            // the "parse" method also validates XML, will throw an exception if misformatted
        } catch (Exception e) {
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_NO_XML_STYLING, "No XML Styling!" +e);
        }


    }

}
