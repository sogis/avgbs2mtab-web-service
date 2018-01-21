package ch.so.agi.avgbs2mtab.main;

import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class Avgbs2mtabMainTest {

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();

    @Test
    public void inputFileExtensionOutsideXtfOrXmlThrowsException() throws Exception {
        File sqlFile = createFileWithoutXTFExtension();
        File outputFilePath = validOutputFilePath();
        try{
            Avgbs2mtabMain.runConversion(sqlFile.getAbsolutePath(), outputFilePath.getAbsolutePath());
        } catch (Avgbs2MtabException e) {
            Assert.assertEquals(Avgbs2MtabException.TYPE_WRONG_EXTENSION, e.getType());
        }

    }

    @Test
    public void withXTFFileWithoutXMLStructureThrowsException() throws Exception {
        File xtfFile = createFileWithXTFExtensionAndNoXMLStructure();
        File outputFilePath = validOutputFilePath();
        try {
            Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());
        } catch (Avgbs2MtabException e) {
            Assert.assertEquals(Avgbs2MtabException.TYPE_NO_XML_STYLING, e.getType());
        }
    }

    @Test
    public void xtfFileOfAnotherModelThrowsException() throws Exception{
        ClassLoader classLoader = getClass().getClassLoader();
        File wrongModellXtf = new File(classLoader.getResource("wrong_Modell.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        try {
            Avgbs2mtabMain.runConversion(wrongModellXtf.getAbsolutePath(), outputFilePath.getAbsolutePath());
        }
        catch (Avgbs2MtabException e) {
            Assert.assertEquals(Avgbs2MtabException.TYPE_TRANSFERDATA_NOT_FOR_AVGBS_MODEL, e.getType());
        }
    }

    @Test
    public void noPermissionToReadXTFFileThrowsException() throws Exception {
        File xtfFile = createNonReadableXTFFile();
        File outputFilePath = validOutputFilePath();
        try {
            Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());
        } catch (Avgbs2MtabException e) {
            Assert.assertEquals(Avgbs2MtabException.TYPE_FILE_NOT_READABLE, e.getType());
        }

    }


    @Test
    public void noPermissionToWriteXLSXFileThrowsException() throws Exception{
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_4002_20150807.xtf").getFile());
        File outputFilePath = validOutputFilePath();
        outputFilePath.getParentFile().setWritable(false);
        try {
            Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());
        } catch (Avgbs2MtabException e){
            Assert.assertEquals(Avgbs2MtabException.TYPE_FOLDER_NOT_WRITEABLE, e.getType());
        }
    }

    /* Not yet implemented - assuming valid file due to small and "expert" user group
    @Ignore
    @Test
    public void xtfFailedValidationThrowsException() throws Exception{
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_4001_20150806_defekt.xtf").getFile());
        try {
            main.runConversion(xtfFile.getAbsolutePath(), "Test.xlsx");
        } catch (Avgbs2MtabException e) {
            Assert.assertEquals("TYPE_VALIDATION_FAILED", e.getType());
        }
    }
    */


    //Parzelle geändert (Zugang), Parzelle gelöscht (an bestehende Parzellen)
    @Test
    public void correctValuesWrittenInExcelAtTransferThreeOldAreasToOneOldParcels() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_4002_20150807.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002407_4002_20150807();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfSO0200002407_4002_20150807();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);
    }

    //Hier stolpere ich über das gleiche - Verstehe From1OldParcel nicht. FromOneOldParcel?
    //Neue Parzelle (Teile von bestehendenParzellen), Parzelle geändert (Teilabgang)
    @Test
    public void correctValuesWrittenInExcelAtNewParcelsFrom1OldParcel() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_4001_20150806.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);


        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002407_4001_20150806();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfSO0200002407_4001_20150806();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);
    }


    @Test
    public void correctValuesWrittenInExcelAtNewDPR() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_40051_20150811.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002407_40051_20150811();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfSO0200002407_40051_20150811();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);
    }

    @Test
    public void correctValuesWrittenInExcelAtDeleteDPR() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_40061_20150814.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002407_40061_20150814();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);

    }


    @Test
    public void correctValuesCalculatedInExcel() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_4004_20150810.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002407_4004_20150810();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfSO0200002407_4004_20150810();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);
        
        Assert.assertTrue(allValuesAreCorrect);
    }


    @Test
    public void correctValuesOfDPRsAndParcelsWrittenInExcel() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002427_809_20170529.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002427_809_20170529();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfSO0200002427_809_20170529();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);

    }

    @Test
    public void worksWithDecimalValuesZerosInDPR() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002578_39732-01_20170908.xml").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfSO0200002578_39732_01_20170908();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfSO0200002578_39732_01_20170908();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);
    }

    @Test
    public void worksWithDecimalValuesElseThanZerosInDPR() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("dprarea_as_double.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfDprarea_as_double();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfDprarea_as_double();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);


        Assert.assertTrue(allValuesAreCorrect);
    }

    @Test
    public void correctDecimalValuesInParcelTable() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("parcelareas_as_double.xtf").getFile());
        File outputFilePath = validOutputFilePath();

        Avgbs2mtabMain.runConversion(xtfFile.getAbsolutePath(), outputFilePath.getAbsolutePath());

        XSSFSheet xlsxSheet = openExcelSheet(outputFilePath.getAbsolutePath());

        HashMap<String, Double> xlsxDataNumeric = generateHashMapFromNumericValuesInExcel(xlsxSheet);
        HashMap<String, String> xlsxDataString = generateHashMapFromStringValuesInExcel(xlsxSheet);

        HashMap<String, String> expectedValuesString =
                generateHashMapOfExpectedStringValuesOfParcelareas_as_double();
        HashMap<String, Double> expectedValuesNumeric =
                generateHashMapOfExpectedNumericValuesOfParcelareas_as_double();

        Boolean allValuesAreCorrect = checkIfValuesAreCorrect(expectedValuesNumeric,
                xlsxDataNumeric,
                expectedValuesString,
                xlsxDataString);

        Assert.assertTrue(allValuesAreCorrect);
    }



    private File createFileWithoutXTFExtension() throws Exception {
        File noXtfExtensionFile =  folder.newFile("query.sql");
        BufferedWriter writer1 = new BufferedWriter(new FileWriter(noXtfExtensionFile));
        writer1.write("Test");
        writer1.close();
        return noXtfExtensionFile ;
    }


    private File createFileWithXTFExtensionAndNoXMLStructure() throws Exception {
        File XtfExtensionFile =  folder.newFile("query.xtf");
        BufferedWriter writer1 = new BufferedWriter(new FileWriter(XtfExtensionFile));
        writer1.write("Test");
        writer1.write("Dies ist ein super Text");
        writer1.close();
        return XtfExtensionFile ;
    }

    private File createNonReadableXTFFile() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File xtfFile = new File(classLoader.getResource("SO0200002407_40053_20150813.xtf").getFile());
        xtfFile.setReadable(false);
        return xtfFile;
    }



    private HashMap<String, String> generateHashMapFromStringValuesInExcel(XSSFSheet xlsxSheet) throws Exception {
        Iterator<Row> rowIterator = xlsxSheet.iterator();

        HashMap<String, String> xlsxDataString = new HashMap<>();

        while (rowIterator.hasNext()) {
            Row row =rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                int rowNumber = row.getRowNum() + 1;

                String key = CellReference.convertNumToColString(cell.getColumnIndex()) + rowNumber;

                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        xlsxDataString.put(key, cell.getStringCellValue());
                        break;
                    default:
                        break;
                }
            }
        }

        return xlsxDataString;
    }


    private HashMap<String, Double> generateHashMapFromNumericValuesInExcel(XSSFSheet xlsxSheet) throws Exception {
        Iterator<Row> rowIterator = xlsxSheet.iterator();

        HashMap<String, Double> xlsxDataNumeric = new HashMap<>();

        while (rowIterator.hasNext()) {
            Row row =rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                int rowNumber = row.getRowNum() + 1;

                String key = CellReference.convertNumToColString(cell.getColumnIndex()) + rowNumber;


                switch (cell.getCellTypeEnum()) {

                    case NUMERIC:
                        xlsxDataNumeric.put(key, cell.getNumericCellValue());
                        break;
                    default:
                        break;
                }
            }
        }
        return xlsxDataNumeric;
    }


    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002407_4002_20150807() {
        HashMap<String, String> expectedValuesString = new HashMap<>();
        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A5", "2199");
        expectedValuesString.put("A7", "Rundungsdifferenz");
        expectedValuesString.put("A8", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");
        expectedValuesString.put("B3", "681");

        expectedValuesString.put("C1", "Alte Liegenschaften");
        expectedValuesString.put("C3", "682");

        expectedValuesString.put("D1", "Alte Liegenschaften");
        expectedValuesString.put("D3", "2199");

        expectedValuesString.put("E2", "Neue Fläche");
        expectedValuesString.put("E3", "[m2]");

        expectedValuesString.put("A12", "Selbst. Recht");
        expectedValuesString.put("A13", "Grundstück-Nr.");

        expectedValuesString.put("B11", "Liegenschaften");
        expectedValuesString.put("B12", "Grundstück-Nr.");

        expectedValuesString.put("C12", "Rundungsdifferenz");

        expectedValuesString.put("D12", "Selbst. Recht Fläche");
        expectedValuesString.put("D13", "[m2]");

        return expectedValuesString;

    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfSO0200002407_4002_20150807() {

        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();


        expectedValuesNumeric.put("B5", (double) 275);
        expectedValuesNumeric.put("B8", (double) 275);


        expectedValuesNumeric.put("C5", (double) 440);
        expectedValuesNumeric.put("C8", (double) 440);


        expectedValuesNumeric.put("D5", (double) 1525);
        expectedValuesNumeric.put("D8", (double) 1525);

        expectedValuesNumeric.put("E5", (double) 2240);
        expectedValuesNumeric.put("E8", (double) 2240);

        return expectedValuesNumeric;
    }



    private boolean checkIfValuesAreCorrect (HashMap<String, Double> expectedValuesNumeric,
                                             HashMap<String, Double> xlsxDataNumeric,
                                             HashMap<String, String> expectedValuesString,
                                             HashMap<String, String> xlsxDataString) throws Exception{
        boolean allCheckedValuesAreCorrect = true;

        allCheckedValuesAreCorrect = checkIfAllExpectedNumericCellsHaveValues(expectedValuesNumeric,
                xlsxDataNumeric,
                allCheckedValuesAreCorrect);

        allCheckedValuesAreCorrect = checkIfAllExpectedStringCellsHaveValues(expectedValuesString,
                xlsxDataString,
                allCheckedValuesAreCorrect);

        allCheckedValuesAreCorrect = checkIfAllNumericValuesAreAsExpected(expectedValuesNumeric,
                xlsxDataNumeric,
                allCheckedValuesAreCorrect);

        allCheckedValuesAreCorrect = checkIfAllStringValuesAreAsExpected(expectedValuesString,
                xlsxDataString,
                allCheckedValuesAreCorrect);

        return allCheckedValuesAreCorrect;
    }

    private boolean checkIfAllExpectedNumericCellsHaveValues(HashMap<String, Double> expectedValuesNumeric,
                                                            HashMap<String, Double> xlsxDataNumeric,
                                                     Boolean allCheckedValuesAreCorrect ) throws Exception {
        for (String key : expectedValuesNumeric.keySet()) {
            if(xlsxDataNumeric.get(key)==null){
                allCheckedValuesAreCorrect = false;
            }

        }
        return allCheckedValuesAreCorrect;
    }


    private boolean checkIfAllExpectedStringCellsHaveValues(HashMap<String, String> expectedValuesString,
                                                            HashMap<String, String> xlsxDataString,
                                                            Boolean allCheckedValuesAreCorrect ) throws Exception {
        for (String key : expectedValuesString.keySet()) {
            if(xlsxDataString.get(key)==null){
                allCheckedValuesAreCorrect = false;
            }
        }
        return allCheckedValuesAreCorrect;
    }


    private boolean checkIfAllNumericValuesAreAsExpected(HashMap<String, Double> expectedValuesNumeric,
                                                         HashMap<String, Double> xlsxDataNumeric,
                                                         Boolean allCheckedValuesAreCorrect) throws Exception{
        for (Map.Entry<String, Double> entry : xlsxDataNumeric.entrySet()) {
            String key = entry.getKey();
            Double value = entry.getValue();

            if (expectedValuesNumeric.get(key)!=null) {
                if (value.compareTo(expectedValuesNumeric.get(key))!=0) {
                    allCheckedValuesAreCorrect = false;
                }
            } else {
                allCheckedValuesAreCorrect = false;
            }
        }

        return allCheckedValuesAreCorrect;
    }

    private boolean checkIfAllStringValuesAreAsExpected(HashMap<String, String> expectedValuesString,
                                                        HashMap<String, String> xlsxDataString,
                                                        Boolean allCheckedValuesAreCorrect) throws Exception {

        for (Map.Entry<String, String> entry : xlsxDataString.entrySet()) {
            String key = entry.getKey();
            String value = entry.getValue();

            if (value != null && !value.equals("")) {
                if (expectedValuesString.get(key) != null) {
                    if (!value.equals(expectedValuesString.get(key))) {
                        allCheckedValuesAreCorrect = false;
                    }
                } else {
                    allCheckedValuesAreCorrect = false;
                }
            }

        }

        return allCheckedValuesAreCorrect;
    }

    private XSSFSheet openExcelSheet(String filePath) {
        try {
            InputStream ExcelFileToRead = new FileInputStream(filePath);
            XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
            return wb.getSheetAt(0);
        } catch (Exception e){
            throw new RuntimeException(e);
        }

    }


    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002407_4001_20150806() {
        HashMap<String, String> expectedValuesString = new HashMap<>();
        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A5", "597");
        expectedValuesString.put("A7", "40011");
        expectedValuesString.put("A9", "40012");
        expectedValuesString.put("A11", "40013");
        expectedValuesString.put("A13", "Rundungsdifferenz");
        expectedValuesString.put("A14", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");
        expectedValuesString.put("B3", "597");

        expectedValuesString.put("C2", "Neue Fläche");
        expectedValuesString.put("C3", "[m2]");

        expectedValuesString.put("A18", "Selbst. Recht");
        expectedValuesString.put("A19", "Grundstück-Nr.");

        expectedValuesString.put("B17", "Liegenschaften");
        expectedValuesString.put("B18", "Grundstück-Nr.");

        expectedValuesString.put("C18", "Rundungs-differenz");

        expectedValuesString.put("D18", "Selbst. Recht Fläche");
        expectedValuesString.put("D19", "[m2]");

        return expectedValuesString;
    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfSO0200002407_4001_20150806() {

        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();



        expectedValuesNumeric.put("B5", (double) 636);
        expectedValuesNumeric.put("B7", (double) 453);
        expectedValuesNumeric.put("B9", (double) 460);
        expectedValuesNumeric.put("B11", (double) 402);
        expectedValuesNumeric.put("B13", (double) -1);
        expectedValuesNumeric.put("B14", (double) 1950);

        expectedValuesNumeric.put("C5", (double) 636);
        expectedValuesNumeric.put("C7", (double) 453);
        expectedValuesNumeric.put("C9", (double) 460);
        expectedValuesNumeric.put("C11", (double) 402);
        expectedValuesNumeric.put("C13", (double) -1);
        expectedValuesNumeric.put("C14", (double) 1950);

        return expectedValuesNumeric;
    }

    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002407_40051_20150811() {
        HashMap<String, String> expectedValuesString = new HashMap<>();
        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A7", "Rundungsdifferenz");
        expectedValuesString.put("A8", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");

        expectedValuesString.put("C2", "Neue Fläche");
        expectedValuesString.put("C3", "[m2]");

        expectedValuesString.put("A12", "Selbst. Recht");
        expectedValuesString.put("A13", "Grundstück-Nr.");
        expectedValuesString.put("A15", "(40051)");

        expectedValuesString.put("B11", "Liegenschaften");
        expectedValuesString.put("B12", "Grundstück-Nr.");
        expectedValuesString.put("B13", "2141");

        expectedValuesString.put("C11", "Liegenschaften");
        expectedValuesString.put("C13", "2142");

        expectedValuesString.put("D12", "Rundungs-differenz");

        expectedValuesString.put("E12", "Selbst. Recht Fläche");
        expectedValuesString.put("E13", "[m2]");

        return expectedValuesString;

    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfSO0200002407_40051_20150811() {

        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();

        expectedValuesNumeric.put("B15", (double) 1175);

        expectedValuesNumeric.put("C15", (double) 2481);
        expectedValuesNumeric.put("E15", (double) 3656);


        return expectedValuesNumeric;
    }


    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002407_40061_20150814() {
        HashMap<String, String> expectedValuesString = new HashMap<>();
        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A7", "Rundungsdifferenz");
        expectedValuesString.put("A8", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");

        expectedValuesString.put("C2", "Neue Fläche");
        expectedValuesString.put("C3", "[m2]");

        expectedValuesString.put("A12", "Selbst. Recht");
        expectedValuesString.put("A13", "Grundstück-Nr.");
        expectedValuesString.put("A15", "(40051)");

        expectedValuesString.put("B11", "Liegenschaften");
        expectedValuesString.put("B12", "Grundstück-Nr.");

        expectedValuesString.put("C12", "Rundungs-differenz");

        expectedValuesString.put("D12", "Selbst. Recht Fläche");
        expectedValuesString.put("D13", "[m2]");
        expectedValuesString.put("D15", "gelöscht");

        return expectedValuesString;

    }

    private boolean checkIfValuesAreCorrect (HashMap<String, String> expectedValuesString,
                                             HashMap<String, String> xlsxDataString) throws Exception{
        boolean allCheckedValuesAreCorrect = true;

        allCheckedValuesAreCorrect = checkIfAllExpectedStringCellsHaveValues(expectedValuesString,
                xlsxDataString,
                allCheckedValuesAreCorrect);

        allCheckedValuesAreCorrect = checkIfAllStringValuesAreAsExpected(expectedValuesString,
                xlsxDataString,
                allCheckedValuesAreCorrect);

        return allCheckedValuesAreCorrect;
    }

    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002407_4004_20150810() {
        HashMap<String, String> expectedValuesString = new HashMap<>();

        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A5", "695");
        expectedValuesString.put("A7", "696");
        expectedValuesString.put("A9", "697");
        expectedValuesString.put("A11", "701");
        expectedValuesString.put("A13", "870");
        expectedValuesString.put("A15", "874");
        expectedValuesString.put("A17", "4004");
        expectedValuesString.put("A19", "Rundungsdifferenz");
        expectedValuesString.put("A20", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");
        expectedValuesString.put("B3", "695");

        expectedValuesString.put("C1", "Alte Liegenschaften");
        expectedValuesString.put("C3", "696");

        expectedValuesString.put("D1", "Alte Liegenschaften");
        expectedValuesString.put("D3", "697");

        expectedValuesString.put("E1", "Alte Liegenschaften");
        expectedValuesString.put("E3", "701");

        expectedValuesString.put("F1", "Alte Liegenschaften");
        expectedValuesString.put("F3", "870");

        expectedValuesString.put("G1", "Alte Liegenschaften");
        expectedValuesString.put("G3", "874");

        expectedValuesString.put("H2", "Neue Fläche");
        expectedValuesString.put("H3", "[m2]");

        expectedValuesString.put("A24", "Selbst. Recht");
        expectedValuesString.put("A25", "Grundstück-Nr.");

        expectedValuesString.put("B23", "Liegenschaften");
        expectedValuesString.put("B24", "Grundstück-Nr.");

        expectedValuesString.put("C24", "Rundungsdifferenz");

        expectedValuesString.put("D24", "Selbst. Recht Fläche");
        expectedValuesString.put("D25", "[m2]");

        return expectedValuesString;

    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfSO0200002407_4004_20150810() {



        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();

        expectedValuesNumeric.put("B5", (double) 416);
        expectedValuesNumeric.put("B17", (double) 242);
        expectedValuesNumeric.put("B19", (double) -1);
        expectedValuesNumeric.put("B20", (double) 657);

        expectedValuesNumeric.put("C7", (double) 507);
        expectedValuesNumeric.put("C11", (double) 1);
        expectedValuesNumeric.put("C17", (double) 100);
        expectedValuesNumeric.put("C20", (double) 608);

        expectedValuesNumeric.put("D9", (double) 687);
        expectedValuesNumeric.put("D11", (double) 1);
        expectedValuesNumeric.put("D17", (double) 129);
        expectedValuesNumeric.put("D19", (double) -1);
        expectedValuesNumeric.put("D20", (double) 816);

        expectedValuesNumeric.put("E11", (double) 1112);
        expectedValuesNumeric.put("E17", (double) 1);
        expectedValuesNumeric.put("E19", (double) 1);
        expectedValuesNumeric.put("E20", (double) 1114);

        expectedValuesNumeric.put("F13", (double) 611);
        expectedValuesNumeric.put("F17", (double) 39);
        expectedValuesNumeric.put("F20", (double) 650);

        expectedValuesNumeric.put("G15", (double) 1939);
        expectedValuesNumeric.put("G17", (double) 81);
        expectedValuesNumeric.put("G20", (double) 2020);

        expectedValuesNumeric.put("H5", (double) 416);
        expectedValuesNumeric.put("H7", (double) 507);
        expectedValuesNumeric.put("H9", (double) 687);
        expectedValuesNumeric.put("H11", (double) 1114);
        expectedValuesNumeric.put("H13", (double) 611);
        expectedValuesNumeric.put("H15", (double) 1939);
        expectedValuesNumeric.put("H17", (double) 592);
        expectedValuesNumeric.put("H19", (double) -1);
        expectedValuesNumeric.put("H20", (double) 5865);


        return expectedValuesNumeric;
    }


    private File validOutputFilePath() {
        File res = null;
        try {
            File subFolder = folder.newFolder();
            String filePath = subFolder.getAbsolutePath() + ".xlsx";
            res = new File(filePath);
        } catch (IOException ioe) {
            throw new Avgbs2MtabException("Could not create testing subfolder in: " + folder.toString());
        }

        return res;
    }

    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002427_809_20170529() {
        HashMap<String, String> expectedValuesString = new HashMap<>();

        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A5", "1273");
        expectedValuesString.put("A7", "1864");
        expectedValuesString.put("A9", "1973");
        expectedValuesString.put("A11", "Rundungsdifferenz");
        expectedValuesString.put("A12", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");
        expectedValuesString.put("B3", "1273");

        expectedValuesString.put("C1", "Alte Liegenschaften");
        expectedValuesString.put("C3", "1864");

        expectedValuesString.put("D2", "Neue Fläche");
        expectedValuesString.put("D3", "[m2]");

        expectedValuesString.put("A16", "Selbst. Recht");
        expectedValuesString.put("A17", "Grundstück-Nr.");
        expectedValuesString.put("A19", "(1941)");

        expectedValuesString.put("B15", "Liegenschaften");
        expectedValuesString.put("B16", "Grundstück-Nr.");
        expectedValuesString.put("B17", "1864");

        expectedValuesString.put("C16", "Rundungsdifferenz");

        expectedValuesString.put("D16", "Selbst. Recht Fläche");
        expectedValuesString.put("D17", "[m2]");

        return expectedValuesString;
    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfSO0200002427_809_20170529() {
        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();



        expectedValuesNumeric.put("B5", (double) 352);
        expectedValuesNumeric.put("B9", (double) 2);
        expectedValuesNumeric.put("B12", (double) 354);


        expectedValuesNumeric.put("C5", (double) 69);
        expectedValuesNumeric.put("C7", (double) 1238);
        expectedValuesNumeric.put("C9", (double) 1434);
        expectedValuesNumeric.put("C12", (double) 2741);

        expectedValuesNumeric.put("D5", (double) 421);
        expectedValuesNumeric.put("D7", (double) 1238);
        expectedValuesNumeric.put("D9", (double) 1436);
        expectedValuesNumeric.put("D12", (double) 3095);

        expectedValuesNumeric.put("B19", (double) 1238);

        expectedValuesNumeric.put("D19", (double) 1238);

        return expectedValuesNumeric;
    }

    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfSO0200002578_39732_01_20170908() {
        HashMap<String, String> expectedValuesString = new HashMap<>();

        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A7", "Rundungsdifferenz");
        expectedValuesString.put("A8", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");

        expectedValuesString.put("C2", "Neue Fläche");
        expectedValuesString.put("C3", "[m2]");

        expectedValuesString.put("A12", "Selbst. Recht");
        expectedValuesString.put("A13", "Grundstück-Nr.");
        expectedValuesString.put("A15", "(1635)");

        expectedValuesString.put("B11", "Liegenschaften");
        expectedValuesString.put("B12", "Grundstück-Nr.");
        expectedValuesString.put("B13", "1606");

        expectedValuesString.put("C12", "Rundungs-differenz");

        expectedValuesString.put("D12", "Selbst. Recht Fläche");
        expectedValuesString.put("D13", "[m2]");

        return expectedValuesString;
    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfSO0200002578_39732_01_20170908() {
        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();

        expectedValuesNumeric.put("B15", (double) 893);

        expectedValuesNumeric.put("D15", (double) 893);

        return expectedValuesNumeric;
    }

    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfDprarea_as_double() {
        HashMap<String, String> expectedValuesString = new HashMap<>();

        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A7", "Rundungsdifferenz");
        expectedValuesString.put("A8", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");

        expectedValuesString.put("C2", "Neue Fläche");
        expectedValuesString.put("C3", "[m2]");

        expectedValuesString.put("A12", "Selbst. Recht");
        expectedValuesString.put("A13", "Grundstück-Nr.");
        expectedValuesString.put("A15", "(1635)");

        expectedValuesString.put("B11", "Liegenschaften");
        expectedValuesString.put("B12", "Grundstück-Nr.");
        expectedValuesString.put("B13", "1606");

        expectedValuesString.put("C12", "Rundungs-differenz");

        expectedValuesString.put("D12", "Selbst. Recht Fläche");
        expectedValuesString.put("D13", "[m2]");

        return expectedValuesString;
    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfDprarea_as_double() {
        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();

        expectedValuesNumeric.put("B15", 123.5);

        expectedValuesNumeric.put("D15", 123.5);

        return expectedValuesNumeric;
    }

    private HashMap<String, String> generateHashMapOfExpectedStringValuesOfParcelareas_as_double() {
        HashMap<String, String> expectedValuesString = new HashMap<>();

        expectedValuesString.put("A2", "Neue Liegenschaften");
        expectedValuesString.put("A3", "Grundstück-Nr.");
        expectedValuesString.put("A5", "695");
        expectedValuesString.put("A7", "696");
        expectedValuesString.put("A9", "697");
        expectedValuesString.put("A11", "701");
        expectedValuesString.put("A13", "870");
        expectedValuesString.put("A15", "874");
        expectedValuesString.put("A17", "4004");
        expectedValuesString.put("A19", "Rundungsdifferenz");
        expectedValuesString.put("A20", "Alte Fläche [m2]");

        expectedValuesString.put("B1", "Alte Liegenschaften");
        expectedValuesString.put("B2", "Grundstück-Nr.");
        expectedValuesString.put("B3", "695");

        expectedValuesString.put("C1", "Alte Liegenschaften");
        expectedValuesString.put("C3", "696");

        expectedValuesString.put("D1", "Alte Liegenschaften");
        expectedValuesString.put("D3", "697");

        expectedValuesString.put("E1", "Alte Liegenschaften");
        expectedValuesString.put("E3", "701");

        expectedValuesString.put("F1", "Alte Liegenschaften");
        expectedValuesString.put("F3", "870");

        expectedValuesString.put("G1", "Alte Liegenschaften");
        expectedValuesString.put("G3", "874");

        expectedValuesString.put("H2", "Neue Fläche");
        expectedValuesString.put("H3", "[m2]");

        expectedValuesString.put("A24", "Selbst. Recht");
        expectedValuesString.put("A25", "Grundstück-Nr.");

        expectedValuesString.put("B23", "Liegenschaften");
        expectedValuesString.put("B24", "Grundstück-Nr.");

        expectedValuesString.put("C24", "Rundungsdifferenz");

        expectedValuesString.put("D24", "Selbst. Recht Fläche");
        expectedValuesString.put("D25", "[m2]");

        return expectedValuesString;

    }

    private HashMap<String, Double> generateHashMapOfExpectedNumericValuesOfParcelareas_as_double() {



        HashMap<String, Double> expectedValuesNumeric = new HashMap<>();

        expectedValuesNumeric.put("B5", (double) 416);
        expectedValuesNumeric.put("B17", (double) 242);
        expectedValuesNumeric.put("B19", (double) -1);
        expectedValuesNumeric.put("B20", (double) 657);

        expectedValuesNumeric.put("C7", (double) 507);
        expectedValuesNumeric.put("C11", (double) 1);
        expectedValuesNumeric.put("C17",  100.5);
        expectedValuesNumeric.put("C20",  608.5);

        expectedValuesNumeric.put("D9", (double) 687);
        expectedValuesNumeric.put("D11", (double) 1);
        expectedValuesNumeric.put("D17", (double) 129);
        expectedValuesNumeric.put("D19", (double) -1);
        expectedValuesNumeric.put("D20", (double) 816);

        expectedValuesNumeric.put("E11", (double) 1112);
        expectedValuesNumeric.put("E17",  1.2);
        expectedValuesNumeric.put("E19", (double) 1);
        expectedValuesNumeric.put("E20",  1114.2);

        expectedValuesNumeric.put("F13", (double) 611);
        expectedValuesNumeric.put("F17", (double) 39);
        expectedValuesNumeric.put("F20", (double) 650);

        expectedValuesNumeric.put("G15",  1939.1);
        expectedValuesNumeric.put("G17", (double) 81);
        expectedValuesNumeric.put("G20",  2020.1);

        expectedValuesNumeric.put("H5", (double) 416);
        expectedValuesNumeric.put("H7", (double) 507);
        expectedValuesNumeric.put("H9", (double) 687);
        expectedValuesNumeric.put("H11", (double) 1114);
        expectedValuesNumeric.put("H13", (double) 611);
        expectedValuesNumeric.put("H15",  1939.1);
        expectedValuesNumeric.put("H17",  592.7);
        expectedValuesNumeric.put("H19", (double) -1);
        expectedValuesNumeric.put("H20",  5865.8);


        return expectedValuesNumeric;
    }


}
