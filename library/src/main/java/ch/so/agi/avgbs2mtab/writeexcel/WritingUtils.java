package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.DataExtractionParcel;
import ch.so.agi.avgbs2mtab.util.Avgbs2MtabException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

class WritingUtils {

    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());

    private static final Integer aParcelOrADprNeedsTwoRows = 2;
    private static final Integer columnIndexOfNewParcelRow = 0;

    private static final String typeDPR = "dpr";
    private static final String typeParcel = "parcel";
    private static final String typeArea = "area";
    private static final String typeRoundingDifference = "rounding difference";

    /**
     * writes inflow or outflow value in a specific cell
     * @param rowIndex      Index of row
     * @param columnIndex   Index of column
     * @param xlsxSheet     excel sheet
     * @param value         value of inflow or outflow
     */
    void writeValueIntoCell(Integer rowIndex,
                                    Integer columnIndex,
                                    XSSFSheet xlsxSheet,
                                    Integer value,
                                    String type){

        Row row = xlsxSheet.getRow(rowIndex);
        Cell cell = row.getCell(columnIndex);
        double valueDouble = 0.0;
        if (type.equals(typeArea)) {
            valueDouble = value / 10.0;
        } else if (type.equals(typeRoundingDifference)){
            valueDouble = value;
        }
        cell.setCellValue(valueDouble);

    }

    /**
     * Rewrites the number of the given dpr and writes it into a specific cell
     * @param rowIndex      index of row
     * @param columnIndex   index of column
     * @param xlsxSheet     excel sheet
     * @param value         number of dpr as a string
     */
    void writeValueIntoCell(int rowIndex,
                            Integer columnIndex,
                            XSSFSheet xlsxSheet,
                            String value,
                            String type){

        XSSFRow row = xlsxSheet.getRow(rowIndex);
        XSSFCell cell =row.getCell(columnIndex);
        if (type.equals(typeDPR))
            cell.setCellValue("(" + value + ")");
        else if (type.equals(typeParcel))
            cell.setCellValue(value);
    }

    /**
     * Gets column index of a specific parcel
     * @param ParcelNumber  number of parcel
     * @param rowNumber     index of row
     * @param xlsxSheet     excel sheet
     * @return              index of column
     */
    int getColumnIndexOfParcelInTable(String ParcelNumber,
                                      int rowNumber,
                                      XSSFSheet xlsxSheet){
        int indexOldParcelNumber = Integer.MIN_VALUE;

        Row row = xlsxSheet.getRow(rowNumber);

        for (Cell cell : row) {
            if (cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().equals(ParcelNumber)) {
                indexOldParcelNumber = cell.getColumnIndex();
                break;
            }
        }

        if(indexOldParcelNumber == Integer.MIN_VALUE){
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not find Parcel " +
                    ParcelNumber);
        }

        return indexOldParcelNumber;
    }

    /**
     * Gets row index of a specific parcel
     * @param newParcelNumber   number of new parcel
     * @param xlsxSheet         excel sheet
     * @return                  index of row
     */
    int getRowIndexOfNewParcelInTable (String newParcelNumber,
                                       XSSFSheet xlsxSheet){

        int indexNewParcelNumber = Integer.MIN_VALUE;

        Iterator<Row> rowIterator = xlsxSheet.iterator();
        rowIterator.next();

        while(rowIterator.hasNext()){
            Row row1 = rowIterator.next();
            Cell cell1 = row1.getCell(columnIndexOfNewParcelRow);

            if (cell1.getCellTypeEnum() == CellType.STRING && cell1.getStringCellValue().equals(newParcelNumber)){
                indexNewParcelNumber = cell1.getRowIndex();
                break;
            }
        }

        if (indexNewParcelNumber == Integer.MIN_VALUE) {
            LOGGER.log(Level.SEVERE, "Could not find parcel " +  newParcelNumber);
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not find Parcel " +
                    newParcelNumber);
        }

        return indexNewParcelNumber;
    }

    /**
     * Gets all new areas of parcels
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param dataExtractionParcel          get-methods for container
     * @return                              list of new areas
     */
    List<Integer> getAllNewAreas (List<String> orderedListOfNewParcelNumbers,
                                  DataExtractionParcel dataExtractionParcel) {

        List<Integer> newAreaList = new ArrayList<>();

        for (String newParcel : orderedListOfNewParcelNumbers) {
            newAreaList.add(dataExtractionParcel.getNewArea(newParcel));
        }

        return newAreaList;
    }


    /**
     * Calculates sum of all rounding differences
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param dataExtractionParcel          get-methods for container
     * @return                              sum of all rounding differences
     */
    Integer calculateRoundingDifference(List<String> orderedListOfOldParcelNumbers,
                                        DataExtractionParcel dataExtractionParcel) {

        Integer roundingDifference = 0;

        for (String oldParcel : orderedListOfOldParcelNumbers) {
            Integer newRoundingDifference = dataExtractionParcel.getRoundingDifference(oldParcel);
            if (newRoundingDifference != null) {
                roundingDifference += newRoundingDifference;
            }
        }

        return - roundingDifference;

    }

    /**
     * Calculates index of parcel row in dpr table
     * @param newParcelNumber   amount of new parcels
     * @param constant          constant value
     * @return                  index of parcel row
     */
    int calculateIndexOfParcelRow (int newParcelNumber,
                                           int constant) {
        int indexOfParcelRow;

        if (newParcelNumber == 0){
            indexOfParcelRow = aParcelOrADprNeedsTwoRows + constant;
        } else {
            indexOfParcelRow = newParcelNumber * aParcelOrADprNeedsTwoRows + constant;
        }

        return indexOfParcelRow;

    }

    /**
     * gets row index of dpr in dpr table
     * @param indexOfParcelRow  index of row with parcel numbers
     * @param dpr               number of dpr
     * @param xlsxSheet         excel sheet
     * @return                  row index of dpr
     */
    Integer getRowIndexOfDPRInTable(int indexOfParcelRow,
                                    String dpr,
                                    XSSFSheet xlsxSheet){

        int lastRow = xlsxSheet.getLastRowNum();
        Integer indexDPR = null;

        for (int i = indexOfParcelRow + 2; i <= lastRow; i++){
            Row row1 = xlsxSheet.getRow(i);
            Cell cell1 = row1.getCell(0);

            if (cell1.getCellTypeEnum() == CellType.STRING){

                String dprNumber = getDPRNumberFromCell(cell1);

                if (dprNumber.equals(dpr)){
                    indexDPR = cell1.getRowIndex();
                    break;
                }
            }
        }

        if (indexDPR != null) {
            return indexDPR;
        } else {
            LOGGER.log(Level.SEVERE,"Could not finde DPR " + dpr + " in DPR-Table.");
            throw new Avgbs2MtabException(Avgbs2MtabException.TYPE_MISSING_PARCEL_IN_EXCEL, "Could not finde DPR " +
                    dpr + " in DPR-Table.");
        }
    }

    /**
     * extracts the dpr number from a string
     * @param cell  excel cell
     * @return      number of dpr
     */
    String getDPRNumberFromCell(Cell cell){
        String dprString = cell.getStringCellValue();
        int dprStringLength = dprString.length();

        return dprString.substring(1, (dprStringLength-1));
    }

    /**
     * gets all old areas and writes them into a hashmap
     * @param orderedListOfNewParcelNumbers list of numbers of new parcels
     * @param orderedListOfOldParcelNumbers list of numbers of old parcels
     * @param dataExtractionParcel          get-methods for parcel container
     * @return                              hashmap with all old areas
     */
    HashMap<String, Integer> getAllOldAreas(List<String> orderedListOfNewParcelNumbers,
                                             List<String> orderedListOfOldParcelNumbers,
                                             DataExtractionParcel dataExtractionParcel) {

        LOGGER.log(Level.FINER, "Calculating for each old parcel the old area.");

        HashMap<String, Integer> oldAreaHashMap = new HashMap<>();
        Integer oldArea;
        Integer area;
        Integer roundingDifference;

        for (String oldParcel : orderedListOfOldParcelNumbers) {
            oldArea = 0;
            for (String newParcel : orderedListOfNewParcelNumbers) {

                area = getAreaOfFlowBetweenOldAndNewParcel(oldParcel, newParcel, dataExtractionParcel);

                if (area != null) {
                    oldArea += area;
                }
            }

            if (oldArea != null) {
                roundingDifference = dataExtractionParcel.getRoundingDifference(oldParcel);
                if (roundingDifference != null){
                    oldArea = oldArea - roundingDifference*10;
                }
                oldAreaHashMap.put(oldParcel, oldArea);
            } else {
                LOGGER.log(Level.SEVERE,"Area of old parcel must not be null");
                throw new Avgbs2MtabException("Area of old parcel must not be null");
            }

        }

        return oldAreaHashMap;

    }

    /**
     * gets area flow between two parcels
     * @param oldParcel             number of old parcel
     * @param newParcel             number of new parcel
     * @param dataExtractionParcel  get-methods for container
     * @return                      area flow
     */
    Integer getAreaOfFlowBetweenOldAndNewParcel(String oldParcel,
                                                String newParcel,
                                                DataExtractionParcel dataExtractionParcel){

        Integer area;

        if (!oldParcel.equals(newParcel)) {
            area = dataExtractionParcel.getAddedArea(newParcel, oldParcel);
        } else {
            area = dataExtractionParcel.getRestAreaOfParcel(oldParcel);
        }


        return area;

    }





}
