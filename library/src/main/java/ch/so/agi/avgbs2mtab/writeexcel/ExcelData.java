package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * The Class ExcelData gets the data from the container and writes it into the prepared exceltemplate
 */
class ExcelData {

    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());
    private static final ParcelTableWriter parcelTableWriter = new ParcelTableWriter();
    private static final DPRTableWriter dprTableWriter = new DPRTableWriter();



    /**
     * gets the parcel data from the container and writes it into the parcel table from the prepared exceltemplate
     * @param filePath                  Path, where the excel-template should be written to
     * @param workbook                  Excel-workbook
     * @param dataExtractionParcel      Methods to get data from container
     */
    void fillValuesIntoParcelTable (String filePath,
                                           XSSFWorkbook workbook,
                                           DataExtractionParcel dataExtractionParcel){


        try {
            OutputStream excelFile = new FileOutputStream(filePath);
            XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

            List<String> orderedListOfOldParcelNumbers = dataExtractionParcel.getOrderedListOfOldParcelNumbers();
            List<String> orderedListOfNewParcelNumbers = dataExtractionParcel.getOrderedListOfNewParcelNumbers();

            parcelTableWriter.writeParcelsIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    xlsxSheet);

            parcelTableWriter.writeAllInflowsAndOutFlowsIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, dataExtractionParcel, xlsxSheet);

            parcelTableWriter.writeAllRoundingDifferenceIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, dataExtractionParcel, xlsxSheet);

            parcelTableWriter.writeSumOfRoundingDifferenceIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, dataExtractionParcel, xlsxSheet);

            parcelTableWriter.writeAllNewAreasIntoParcelTable(orderedListOfNewParcelNumbers, dataExtractionParcel,
                    xlsxSheet);

            parcelTableWriter.writeAllOldAreasIntoParcelTable(orderedListOfOldParcelNumbers,
                    orderedListOfNewParcelNumbers, dataExtractionParcel, xlsxSheet);

            parcelTableWriter.writeAreaSumIntoParcelTable(orderedListOfOldParcelNumbers, orderedListOfNewParcelNumbers,
                    dataExtractionParcel, xlsxSheet);

            workbook.write(excelFile);
            excelFile.close();

        } catch (IOException e){
            LOGGER.log(Level.SEVERE, "Could not write values into parcel table " + e.getMessage());
            throw new RuntimeException(e);
        }

    }



    /**
     * Fills all values (parcel and dpr number, area values, rounding differences) into dpr table
     * @param filePath                  path, where the excel file should be saved to
     * @param workbook                  excel workbook
     * @param dataExtractionDPR         get-Methods for dpr-container
     * @param metadataOfParcelMutation  get-Methods for dpr metadata
     */
    void fillValuesIntoDPRTable (String filePath, XSSFWorkbook workbook,
                                        DataExtractionDPR dataExtractionDPR,
                                        MetadataOfParcelMutation metadataOfParcelMutation) {

        try {
            OutputStream ExcelFile = new FileOutputStream(filePath);
            XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

            List<String> orderedListOfParcelNumbers = dataExtractionDPR.getOrderedListOfParcelsAffectedByDPRs();
            List<String> orderedListOfDPRs = dataExtractionDPR.getOrderedListOfNewDPRs();


            Integer numberOfNewParcelsInParcelTable = metadataOfParcelMutation.getNumberOfNewParcels();

            dprTableWriter.writeParcelsAndDPRsIntoTable(orderedListOfParcelNumbers, orderedListOfDPRs,
                    numberOfNewParcelsInParcelTable, xlsxSheet);


            dprTableWriter.writeAllFlowsIntoDPRTable(orderedListOfParcelNumbers, orderedListOfDPRs,
                    numberOfNewParcelsInParcelTable, dataExtractionDPR, xlsxSheet);


            dprTableWriter.writeAllRoundingDifferencesIntoDPRTable(orderedListOfDPRs, numberOfNewParcelsInParcelTable,
                    dataExtractionDPR, xlsxSheet);


            dprTableWriter.writeAllNewAreasIntoDPRTable(orderedListOfDPRs, numberOfNewParcelsInParcelTable,
                    dataExtractionDPR, xlsxSheet);


            workbook.write(ExcelFile);
            ExcelFile.close();

        } catch (IOException e){
            LOGGER.log(Level.SEVERE, "Could not write values into dpr table : " + e.getMessage());
            throw new RuntimeException(e);
        }
    }


}
