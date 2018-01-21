package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.DataExtractionParcel;
import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

class ParcelTableWriter {
    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());

    private static final String typeParcel = "parcel";
    private static final String typeArea = "area";
    private static final String typeRoundingDifference = "rounding difference";

    private static final Integer startIndexOfColumn = 1;
    private static final Integer rowIndexOfOldParcels = 2;
    private static final Integer aParcelOrADprNeedsTwoRows = 2;
    private static final Integer startIndexOfRow = 2;
    private static final Integer columnIndexOfNewParcelRow = 0;
    private static final Integer additionConstantToGetToRoundingDifferenceRow = 4; //empty, new Area, square Meter, 2
    // cells rounding difference
    // (-1 because index starts with 0)
    private static final Integer additionConstantToGetToOldAreaRow = additionConstantToGetToRoundingDifferenceRow + 1;
    private static final WritingUtils writingUtils = new WritingUtils();




    /**
     * Writes numbers of parcels into parcel table
     * @param orderedListOfOldParcelNumbers List of old parcel numbers
     * @param orderedListOfNewParcelNumbers List of new parcel numbers
     * @param xlsxSheet                     excel sheet
     */
    void writeParcelsIntoParcelTable(List<String> orderedListOfOldParcelNumbers,
                                             List<String> orderedListOfNewParcelNumbers,
                                             XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write number of new and old parcels into parcel table");

        writeOldParcelsInTemplate(orderedListOfOldParcelNumbers, xlsxSheet);
        writeNewParcelsInTemplate(orderedListOfNewParcelNumbers, xlsxSheet);
    }

    /**
     * writes the numbers of old parcels into specific row of parcel table
     * @param orderedListOfOldParcelNumbers List of old parcel numbers
     * @param xlsxSheet                     excel sheet
     */
    void writeOldParcelsInTemplate(List<String> orderedListOfOldParcelNumbers,
                                   XSSFSheet xlsxSheet) {

        Row rowWithOldParcelNumbers =xlsxSheet.getRow(rowIndexOfOldParcels);

        Integer column = startIndexOfColumn;

        for (String parcelNumber : orderedListOfOldParcelNumbers){
            Cell cell = rowWithOldParcelNumbers.getCell(column);
            cell.setCellValue(parcelNumber);
            column++;
        }
    }

    /**
     * writes the numbers of new parcels into specific column of parcel table
     * @param orderedListOfNewParcelNumbers List of new parcel numbers
     * @param xlsxSheet                     excel sheet
     */
    void writeNewParcelsInTemplate(List<String> orderedListOfNewParcelNumbers,
                                   XSSFSheet xlsxSheet){


        int amountOfParcels = 1;

        for (String parcelNumber : orderedListOfNewParcelNumbers){
            writingUtils.writeValueIntoCell(startIndexOfRow + aParcelOrADprNeedsTwoRows * amountOfParcels,
                    columnIndexOfNewParcelRow, xlsxSheet, parcelNumber, typeParcel);
            amountOfParcels++;
        }
    }




    /**
     * writes all inflows and outflows into parcel table
     * @param orderedListOfOldParcelNumbers List of old parcel numbers
     * @param orderedListOfNewParcelNumbers List of new parcel numbers
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    void writeAllInflowsAndOutFlowsIntoParcelTable (List<String> orderedListOfOldParcelNumbers,
                                                            List<String> orderedListOfNewParcelNumbers,
                                                            DataExtractionParcel dataExtractionParcel,
                                                            XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write all inflows and outflows of each parcel into parcel table.");


        for (String oldParcel : orderedListOfOldParcelNumbers) {
            for (String newParcel : orderedListOfNewParcelNumbers) {

                Integer area = writingUtils.getAreaOfFlowBetweenOldAndNewParcel(oldParcel, newParcel, dataExtractionParcel);

                if (area != null) {
                    writeInflowAndOutflowOfOneParcelPair(oldParcel, newParcel, area, xlsxSheet);
                }
            }
        }
    }



    /**
     * Writes area flow into parcel table
     * @param oldParcelNumber   number of old parcel
     * @param newParcelNumber   number of new parcel
     * @param area              area flow
     * @param xlsxSheet         excel sheet
     */
    void writeInflowAndOutflowOfOneParcelPair(String oldParcelNumber,
                                                     String newParcelNumber,
                                                     int area,
                                                     XSSFSheet xlsxSheet) {

        int indexOldParcelNumber;
        int indexNewParcelNumber;

        indexOldParcelNumber = writingUtils.getColumnIndexOfParcelInTable(oldParcelNumber, rowIndexOfOldParcels, xlsxSheet);
        indexNewParcelNumber = writingUtils.getRowIndexOfNewParcelInTable(newParcelNumber, xlsxSheet);

        writingUtils.writeValueIntoCell(indexNewParcelNumber, indexOldParcelNumber, xlsxSheet, area, typeArea);
    }

    /**
     * Writes all rounding differences of the parcel table into parcel table
     * @param orderedListOfOldParcelNumbers List of numbers of old parcels
     * @param orderedListOfNewParcelNumbers List of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    void writeAllRoundingDifferenceIntoParcelTable(List<String> orderedListOfOldParcelNumbers,
                                                           List<String> orderedListOfNewParcelNumbers,
                                                           DataExtractionParcel dataExtractionParcel,
                                                           XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the rounding difference for each parcel into parcel table.");

        for (String oldParcel : orderedListOfOldParcelNumbers) {

            Integer roundingDifference =  dataExtractionParcel.getRoundingDifference(oldParcel);

            if (roundingDifference != null && roundingDifference != 0) {

                writeRoundingDifference(oldParcel, -roundingDifference, orderedListOfNewParcelNumbers.size(), xlsxSheet);
            }
        }
    }

    /**
     * Writes a rounding difference from one parcel into parcel table
     * @param oldParcelNumber       number of old parcel
     * @param roundingDifference    value of rounding difference
     * @param numberOfNewParcels    amount of new parcels
     * @param xlsxSheet             excel sheet
     */
    void writeRoundingDifference(String oldParcelNumber,
                                 int roundingDifference,
                                 int numberOfNewParcels,
                                 XSSFSheet xlsxSheet) {

        Integer columnOldParcelNumber = writingUtils.getColumnIndexOfParcelInTable(oldParcelNumber, rowIndexOfOldParcels , xlsxSheet);

        Integer rowOldParcelNumber = additionConstantToGetToRoundingDifferenceRow + aParcelOrADprNeedsTwoRows *
                numberOfNewParcels;

        if (columnOldParcelNumber==null){
            String errorMessage = "The old parcel "+ oldParcelNumber + " could not be found in the excel.";
            LOGGER.log(Level.SEVERE, errorMessage);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, errorMessage);
        } else {
            writingUtils.writeValueIntoCell(rowOldParcelNumber, columnOldParcelNumber, xlsxSheet, roundingDifference,
                    typeRoundingDifference);
        }
    }


    /**
     * Writes the sum of all rounding differences of parcel table into sum cell
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    void writeSumOfRoundingDifferenceIntoParcelTable(List<String> orderedListOfOldParcelNumbers,
                                                             List<String> orderedListOfNewParcelNumbers,
                                                             DataExtractionParcel dataExtractionParcel,
                                                             XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the sum of all rounding differences into parcel table");

        if (orderedListOfNewParcelNumbers.size() != 0 && orderedListOfOldParcelNumbers.size() != 0) {
            writeSumOfRoundingDifference(orderedListOfNewParcelNumbers.size(),
                    orderedListOfOldParcelNumbers.size(),
                    writingUtils.calculateRoundingDifference(orderedListOfOldParcelNumbers, dataExtractionParcel),
                    xlsxSheet);
        }
    }


    /**
     * Writes sum of rounding difference into specific cell
     * @param NumberOfNewParcels        amount of new parcels
     * @param NumberOfOldParcels        amount of old parcels
     * @param roundingDifferenceSum     value of rounding difference
     * @param xlsxSheet                 excel sheet
     */
    private void writeSumOfRoundingDifference (int NumberOfNewParcels,
                                              int NumberOfOldParcels,
                                              int roundingDifferenceSum,
                                              XSSFSheet xlsxSheet){

        int rowNumber = additionConstantToGetToRoundingDifferenceRow + aParcelOrADprNeedsTwoRows * NumberOfNewParcels;
        int columnNumber = NumberOfOldParcels + 1;

        if (roundingDifferenceSum != 0) {
            writingUtils.writeValueIntoCell(rowNumber, columnNumber, xlsxSheet, roundingDifferenceSum,
                    typeRoundingDifference);
        }
    }

    /**
     * Writes all new ares into parcel table
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    void writeAllNewAreasIntoParcelTable(List<String> orderedListOfNewParcelNumbers,
                                                 DataExtractionParcel dataExtractionParcel,
                                                 XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write for each parcel the new area into parcel table.");

        for (String newParcel : orderedListOfNewParcelNumbers){

            int newArea = dataExtractionParcel.getNewArea(newParcel);

            writeNewArea(newParcel, newArea, xlsxSheet);
        }
    }

    /**
     * writes the new area of a parcel into parcel table
     * @param newParcelNumber   number of new parcel
     * @param area              value of area
     * @param xlsxSheet         excel sheet
     */
    void writeNewArea(String newParcelNumber,
                             int area,
                             XSSFSheet xlsxSheet) {

        int rowNewParcelNumber = writingUtils.getRowIndexOfNewParcelInTable(newParcelNumber, xlsxSheet);

        try {
            Row row = xlsxSheet.getRow(rowNewParcelNumber);
            Integer columnNewParcelNumber = row.getLastCellNum()-1;
            writingUtils.writeValueIntoCell(rowNewParcelNumber, columnNewParcelNumber, xlsxSheet, area, typeArea);
        } catch (Exception e){
            LOGGER.log(Level.SEVERE,"Last row could not be found");
            throw new Avgbs2MtabException("Could not find last row");
        }
    }



    /**
     * Writes all old areas into parcel table
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          DataExtractionParcel
     * @param xlsxSheet                     excel sheet
     */
    void writeAllOldAreasIntoParcelTable(List<String> orderedListOfOldParcelNumbers,
                                                 List<String> orderedListOfNewParcelNumbers,
                                                 DataExtractionParcel dataExtractionParcel,
                                                 XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER,"Write for each old parcel the old area into parcel table.");


        HashMap<String, Integer> oldAreaHashMap = writingUtils.getAllOldAreas(orderedListOfNewParcelNumbers,
                orderedListOfOldParcelNumbers, dataExtractionParcel);

        for (String oldParcel : orderedListOfOldParcelNumbers) {
            Integer oldArea = oldAreaHashMap.get(oldParcel);

            writeOldArea(oldParcel, oldArea, orderedListOfNewParcelNumbers.size(), xlsxSheet);
        }
    }

    /**
     * Writes area of a old parcel into parcel table
     * @param oldParcelNumber       number of old parcel
     * @param oldArea               area of old parcel
     * @param numberOfnewParcels    amount of new parcels
     * @param xlsxSheet             excel sheet
     */
    void writeOldArea(String oldParcelNumber,
                      int oldArea,
                      int numberOfnewParcels,
                      XSSFSheet xlsxSheet){

        Integer columnOldParcelNumber = writingUtils.getColumnIndexOfParcelInTable(oldParcelNumber, rowIndexOfOldParcels, xlsxSheet);

        Integer rowOldParcelArea = additionConstantToGetToOldAreaRow + aParcelOrADprNeedsTwoRows * numberOfnewParcels;

        if (columnOldParcelNumber != null){
            writingUtils.writeValueIntoCell(rowOldParcelArea, columnOldParcelNumber, xlsxSheet, oldArea, typeArea);
        } else {
            String errorMessage = "The old parcel " + oldParcelNumber + " could not be found in the excel.";
            LOGGER.log(Level.SEVERE, errorMessage);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, errorMessage);
        }
    }


    /**
     * gets the sum of all area (old or new) and writes it into parcel table
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @param xlsxSheet                     excel sheet
     */
    void writeAreaSumIntoParcelTable(List<String> orderedListOfOldParcelNumbers,
                                             List<String> orderedListOfNewParcelNumbers,
                                             DataExtractionParcel dataExtractionParcel,
                                             XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the sum of the areas into parcel table.");

        HashMap<String, Integer> oldAreaHashMap = writingUtils.getAllOldAreas(orderedListOfNewParcelNumbers,
                orderedListOfOldParcelNumbers, dataExtractionParcel);

        if (orderedListOfNewParcelNumbers.size() != 0 && orderedListOfOldParcelNumbers.size() != 0) {

            writeAreaSum(oldAreaHashMap, writingUtils.getAllNewAreas(orderedListOfNewParcelNumbers, dataExtractionParcel),
                    writingUtils.calculateRoundingDifference(orderedListOfOldParcelNumbers, dataExtractionParcel), xlsxSheet);
        }
    }

    /**
     * Calculates the sum of all areas (old or new)
     * @param oldAreas              hashmap with all old areas
     * @param newAreas              list with all new areas
     * @param roundingDifference    sum of all rounding differences
     * @param xlsxSheet             excel sheet
     */
    void writeAreaSum(HashMap<String, Integer> oldAreas,
                      List<Integer> newAreas,
                      int roundingDifference,
                      XSSFSheet xlsxSheet){

        Integer sumOldAreas = 0;
        Integer sumNewAreas = 0;

        Integer numberOfOldParcels = oldAreas.size();
        Integer numberOfNewParcels = newAreas.size();

        Integer columnIndexOfNewAreas = numberOfOldParcels + 1;


        for (Map.Entry<String, Integer> entry : oldAreas.entrySet()){
            sumOldAreas += entry.getValue();
        }

        for (int area : newAreas){
            sumNewAreas += area;
        }
        sumNewAreas = sumNewAreas + roundingDifference*10;

        if (sumOldAreas.equals( sumNewAreas)){
            writingUtils.writeValueIntoCell(additionConstantToGetToOldAreaRow + aParcelOrADprNeedsTwoRows * numberOfNewParcels,
                    columnIndexOfNewAreas, xlsxSheet,
                    sumOldAreas, typeArea);
        } else {
            LOGGER.log(Level.SEVERE, "The sum of the old areas is not equal to the sum of the new areas.");
            throw new Avgbs2MtabException("The sum of the old areas is not equal to the sum of the new areas.");
        }
    }

}
