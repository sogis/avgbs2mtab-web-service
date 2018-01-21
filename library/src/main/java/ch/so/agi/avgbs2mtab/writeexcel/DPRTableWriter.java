package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.DataExtractionDPR;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

class DPRTableWriter {
    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());

    private static final String typeDPR = "dpr";
    private static final String typeParcel = "parcel";
    private static final String typeArea = "area";
    private static final String typeRoundingDifference = "rounding difference";

    private static final Integer additionConstantToGetToRoundingDifferenceRow = 4; //empty, new Area, square Meter, 2
    // cells rounding difference
    // (-1 because index starts with 0)
    private static final Integer additionConstantToGetToOldAreaRow = additionConstantToGetToRoundingDifferenceRow + 1;
    private static final Integer additionConstantToGetToParcelNumberRowInDPRTable =
            additionConstantToGetToOldAreaRow + 5;
    private static final Integer additionConstantToGetToFirstRowWithDPRNumber =
            additionConstantToGetToParcelNumberRowInDPRTable + 2;
    private static final WritingUtils writingUtils = new WritingUtils();

    /**
     * writes all number of parcels and dprs into dpr table
     * @param orderedListOfParcelNumbers        list of numbers of parcels
     * @param orderedListOfDPRs                 list of number of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param xlsxSheet                         excel sheet
     */
    void writeParcelsAndDPRsIntoTable(List<String> orderedListOfParcelNumbers,
                                      List<String> orderedListOfDPRs,
                                      int numberOfNewParcelsInParcelTable,
                                      XSSFSheet xlsxSheet) {

        LOGGER.log(Level.FINER, "Write the numbers of dprs and parcels into dpr table");

        writeParcelsAffectedByDPRsInTemplate(orderedListOfParcelNumbers, numberOfNewParcelsInParcelTable, xlsxSheet);
        writeDPRsInTemplate(orderedListOfDPRs, numberOfNewParcelsInParcelTable, xlsxSheet);
    }

    /**
     * Writes number of parcels into dpr table
     * @param orderedListOfParcelNumbersAffectedByDPRs  list of numbers of parcel
     * @param newParcelNumber                           amount of new parcels in parcel table
     * @param xlsxSheet                                 excel sheet
     */
    void writeParcelsAffectedByDPRsInTemplate(List<String> orderedListOfParcelNumbersAffectedByDPRs,
                                              int newParcelNumber,
                                              XSSFSheet xlsxSheet) {

        Integer column = 1;

        Integer indexOfParcelRow =
                writingUtils.calculateIndexOfParcelRow(newParcelNumber, additionConstantToGetToParcelNumberRowInDPRTable);


        for (String parcelNumber : orderedListOfParcelNumbersAffectedByDPRs){

            writingUtils.writeValueIntoCell(indexOfParcelRow, column, xlsxSheet, parcelNumber, typeParcel);

            column++;
        }
    }

    /**
     * Writes all dpr into dpr table
     * @param orderedListOfDPRs list of numbers of dprs
     * @param newParcelNumber   amount of new parcels
     * @param xlsxSheet         excel sheet
     */
    void writeDPRsInTemplate(List<String> orderedListOfDPRs,
                             int newParcelNumber,
                             XSSFSheet xlsxSheet) {

        Integer rowIndex = writingUtils.calculateIndexOfParcelRow(newParcelNumber, additionConstantToGetToFirstRowWithDPRNumber);

        for (String dpr : orderedListOfDPRs){

            writingUtils.writeValueIntoCell(rowIndex, 0, xlsxSheet, dpr.toString(), typeDPR);

            rowIndex++;
            rowIndex++;
        }
    }

    /**
     * Writes all flows between parcels and dprs into dpr table
     * @param orderedListOfParcelNumbers        list of numbers of parcels
     * @param orderedListOfDPRs                 list of numbers of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param dataExtractionDPR                 get-methods for dpr container
     * @param xlsxSheet                         excel sheet
     */
    void writeAllFlowsIntoDPRTable(List<String> orderedListOfParcelNumbers,
                                   List<String> orderedListOfDPRs,
                                   int numberOfNewParcelsInParcelTable,
                                   DataExtractionDPR dataExtractionDPR,
                                   XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write all flows of area into dpr table");

        for (String parcel : orderedListOfParcelNumbers) {
            for (String dpr : orderedListOfDPRs) {

                Integer area = dataExtractionDPR.getAddedAreaDPR(parcel, dpr);

                writeDPRInflowAndOutflows(parcel, dpr, area, numberOfNewParcelsInParcelTable, xlsxSheet);
            }
        }
    }

    /**
     * Writes area flow between one parcel and one dpr into dpr table
     * @param parcelNumberAffectedByDPR number of parcel
     * @param dpr                       number of dpr
     * @param area                      value of area
     * @param newParcelNumber           amount of new parcels
     * @param xlsxSheet                 excel sheet
     */
    void writeDPRInflowAndOutflows(String parcelNumberAffectedByDPR,
                                   String dpr,
                                   int area,
                                   int newParcelNumber,
                                   XSSFSheet xlsxSheet) {

        Integer indexOfParcelRow =
                writingUtils.calculateIndexOfParcelRow(newParcelNumber, additionConstantToGetToParcelNumberRowInDPRTable);

        int indexParcel = writingUtils.getColumnIndexOfParcelInTable(parcelNumberAffectedByDPR, indexOfParcelRow, xlsxSheet);

        int indexDPR = writingUtils.getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        writingUtils.writeValueIntoCell(indexDPR, indexParcel, xlsxSheet, area, typeArea);
    }


    /**
     * Writes all rounding differences into dpr table
     * @param orderedListOfDPRs                 list of numbers of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param dataExtractionDPR                 get-Methods for dpr container
     * @param xlsxSheet                         excel sheet
     */
    void writeAllRoundingDifferencesIntoDPRTable(List<String> orderedListOfDPRs,
                                                 int numberOfNewParcelsInParcelTable,
                                                 DataExtractionDPR dataExtractionDPR,
                                                 XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write rounding difference for each dpr into dpr table");

        for (String dpr : orderedListOfDPRs) {

            Integer roundingDifference = dataExtractionDPR.getRoundingDifferenceDPR(dpr);

            if (roundingDifference != null && roundingDifference != 0) {
                writeDPRRoundingDifference(dpr, -roundingDifference, numberOfNewParcelsInParcelTable, xlsxSheet);
            }
        }
    }

    /**
     * Write rounding difference for one dpr
     * @param dpr                   number of dpr
     * @param roundingDifference    rounding difference
     * @param newParcelNumber       amount of new parcels in parcel table
     * @param xlsxSheet             excel sheet
     */
    void writeDPRRoundingDifference(String dpr,
                                    int roundingDifference,
                                    int newParcelNumber,
                                    XSSFSheet xlsxSheet) {

        Integer rowDPRNumber;
        Integer columnRoundingDifference;

        Integer indexOfParcelRow =
                writingUtils.calculateIndexOfParcelRow(newParcelNumber, additionConstantToGetToParcelNumberRowInDPRTable);

        rowDPRNumber = writingUtils.getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        columnRoundingDifference = (int) xlsxSheet.getRow(rowDPRNumber).getLastCellNum()-2;

        writingUtils.writeValueIntoCell(rowDPRNumber, columnRoundingDifference, xlsxSheet, roundingDifference,
                typeRoundingDifference);
    }

    /**
     * writes all new areas into dpr table
     * @param orderedListOfDPRs                 list of numbers of dprs
     * @param numberOfNewParcelsInParcelTable   amount of new parcels in parcel table
     * @param dataExtractionDPR                 get-methods for dpr container
     * @param xlsxSheet                         excel sheet
     */
    void writeAllNewAreasIntoDPRTable(List<String> orderedListOfDPRs,
                                      int numberOfNewParcelsInParcelTable,
                                      DataExtractionDPR dataExtractionDPR,
                                      XSSFSheet xlsxSheet){

        LOGGER.log(Level.FINER, "Write for each dpr the new area into dpr table");

        for (String dpr : orderedListOfDPRs) {
            Integer newArea = dataExtractionDPR.getNewAreaDPR(dpr);
            writeNewDPRArea(dpr, newArea, numberOfNewParcelsInParcelTable, xlsxSheet);
        }
    }


    /**
     * writes the new area of one dpr into dpr table
     * @param dpr               number of dpr
     * @param area              value of area
     * @param newParcelNumber   amount of new parcels in parcel table
     * @param xlsxSheet         excel sheet
     */
    void writeNewDPRArea(String dpr,
                         int area,
                         int newParcelNumber,
                         XSSFSheet xlsxSheet) {

        Integer indexOfParcelRow =
                writingUtils.calculateIndexOfParcelRow(newParcelNumber, additionConstantToGetToParcelNumberRowInDPRTable);

        Integer rowDPRNumber = writingUtils.getRowIndexOfDPRInTable(indexOfParcelRow, dpr, xlsxSheet);

        Integer columnNewArea = (int) xlsxSheet.getRow(rowDPRNumber).getLastCellNum()-1;

        writeNewDPRAreaValueIntoCell(rowDPRNumber, columnNewArea, area, xlsxSheet);
    }

    /**
     * Writes new area value from a specific dpr into dpr table
     * @param rowDPRNumber      index of row of specific dpr
     * @param columnNewArea     index of column with new areas
     * @param area              value of area
     * @param xlsxSheet         excel sheet
     */
    private void writeNewDPRAreaValueIntoCell(int rowDPRNumber,
                                              int columnNewArea,
                                              int area,
                                              XSSFSheet xlsxSheet){

        XSSFRow rowFlows = xlsxSheet.getRow(rowDPRNumber);
        XSSFCell cellFlows = rowFlows.getCell(columnNewArea);
        double areaDouble = area/10.0;

        if (areaDouble > 0) {
            cellFlows.setCellValue(areaDouble);
        } else {
            cellFlows.setCellValue("gelöscht");
        }
    }
}
