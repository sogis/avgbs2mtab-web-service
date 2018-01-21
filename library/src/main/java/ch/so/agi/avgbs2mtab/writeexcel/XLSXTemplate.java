package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.MetadataOfDPRMutation;
import ch.so.agi.avgbs2mtab.mutdat.MetadataOfParcelMutation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;


import java.io.*;
import java.util.logging.Logger;
import java.util.logging.Level;


/**
 * The class XLSXTemplate generates an excel template, where the two tables (parcel and dpr table) are already styled
 */
public class XLSXTemplate implements ExcelTemplate {

    private static final Logger LOGGER = Logger.getLogger( XLSXTemplate.class.getName());

    private static final String xlsxSheetName = "Mutationstabelle";
    private static final String thick ="thick";
    private static final String thin ="thin";
    private static final String lightGrayString = "lightGray";
    private static final String noStyling = "";
    private static final String fontName = "Arial";
    private static final String parcelString = "Liegenschaften";
    private static final String oldParcelString = "Alte " + parcelString;
    private static final String newParcelString = "Neue " + parcelString;
    private static final String parcelNumberString = "Grundstück-Nr.";
    private static final String newAreaString = "Neue Fläche";
    private static final String squareMeterString = "[m2]";
    private static final String oldAreaString = "Alte Fläche " + squareMeterString;
    private static final String roundingDifferenceString = "Rundungsdifferenz";
    private static final String splittedRoundingDifferenceString = "Rundungs-differenz";
    private static final String dprString = "Selbst. Recht";
    private static final String dprAreaString = dprString + " Fläche";

    private static final Integer aParcelOrADprNeedsTwoRows = 2;
    private static final Integer rowsBesideParcelRows = 5;
    private static final Integer widthMultiplicationFactor = 253;
    private static final Integer widthOfFirstColumn = 19 * widthMultiplicationFactor;
    private static final Integer widthOfLastParcelTableColumnAndTheFollowingColumn = 14 * widthMultiplicationFactor;
    private static final Integer indentValue = 2;
    private static final Integer noIndentValue = 0;
    private static final Integer additionConstantToGetIndexOfLastColumnOfParcelTable = 1;
    private static final Integer additionConstantToGetIndexOfRoundingDifferenceColumnOfDPRTable = 1;
    private static final Integer additionConstantToGetIndexOfNewAreaColumnOfDPRTable = 2;

    private static final short rowHeightOfTableHeaderAndFooter = 600;
    private static final short defaultRowHeight = 300;
    private static final short defaultColumnWidth = (short) 18.43;
    private static final short fontHeight = 11;

    private String color;
    private String border_bottom;
    private String border_top;
    private String border_left;
    private String border_right;
    private Integer indent;

    private Cell cell;

    /**
     *
     * Gets metadata and creates an excel template, with two styled tables
     * @param filePath                  path, where excel file should be saved to
     * @param metadataOfParcelMutation  MetadataOfParcelMutation
     * @param metadataOfDPRMutation     MetadataOfDPRMutation
     * @return                          excel template, with two styled tables
     */
    XSSFWorkbook createExcelTemplate (String filePath,MetadataOfParcelMutation metadataOfParcelMutation,
                                             MetadataOfDPRMutation metadataOfDPRMutation ) {

        Integer numberOfNewParcels = metadataOfParcelMutation.getNumberOfNewParcels();
        Integer numberOfOldParcels = metadataOfParcelMutation.getNumberOfOldParcels();

        Integer numberOfParcelsAffectedByDPRs = metadataOfDPRMutation.getNumberOfParcelsAffectedByDPRs();
        Integer numberOfDPRs = metadataOfDPRMutation.getNumberOfDPRs();

        return generateWorkbookTemplate(filePath, numberOfNewParcels, numberOfOldParcels,
                numberOfParcelsAffectedByDPRs, numberOfDPRs);
    }

    /**
     * generates an excel template, with two styled tables
     * @param filePath                          path, where excel file should be saved to
     * @param numberOfNewParcels                amount of new parcels
     * @param numberOfOldParcels                amount of old parcels
     * @param numberOfParcelsAffectedByDPRs     amount of parcels which ar affected by dprs
     * @param numberOfDPRs                      amount of dprs
     * @return                                  excel template with two styled tables
     */
    private XSSFWorkbook generateWorkbookTemplate(String filePath, int numberOfNewParcels, int numberOfOldParcels,
                                                  int numberOfParcelsAffectedByDPRs, int numberOfDPRs){

        LOGGER.log(Level.FINER, "Start creating Excel-Workbook");
        XSSFWorkbook workbook = createWorkbook(filePath);
        LOGGER.log(Level.FINER, "Finished creating Excel-Workbook; Start creating empty table" +
                "with parcels");
        createParcelTable(workbook, filePath, numberOfNewParcels, numberOfOldParcels,
                numberOfParcelsAffectedByDPRs);
        LOGGER.log(Level.FINER, "Finished creating table with parcels; Start creating empty " +
                "table with dprs");
        createDPRTable(workbook, filePath, numberOfParcelsAffectedByDPRs, numberOfDPRs,
                numberOfNewParcels, numberOfOldParcels);
        LOGGER.log(Level.FINER, "Finished creating table with dprs");

        return workbook;
    }


    /**
     * creates an excel workbook and saves it
     * @param filePath  path, where excel file should be saved to
     * @return          excel workbook
     */
    @Override
    public XSSFWorkbook createWorkbook(String filePath) {

        try {
            OutputStream ExcelFile = new FileOutputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook();
            workbook.createSheet(xlsxSheetName);
            workbook.write(ExcelFile);
            ExcelFile.close();

            return workbook;

        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "Could not create Excel-Workbook at " + filePath);
            throw new RuntimeException(e);
        }
    }

    /**
     * creates the parcel table
     * @param excelTemplate         excel workbook
     * @param filePath              path, where excel file should be saved to
     * @param newParcels            amount of new parcels
     * @param oldParcels            amount of old parcels
     * @param parcelsAffectedByDPR  amount of parcels in dpr table
     */
    @Override
    public void createParcelTable(XSSFWorkbook excelTemplate,String filePath, int newParcels, int oldParcels,
                                          int parcelsAffectedByDPR) {

        XSSFSheet sheet = excelTemplate.getSheet(xlsxSheetName);

        mergeCellsToRegions(sheet, oldParcels);

        if (oldParcels == 0 || newParcels == 0){
            oldParcels = 1;
            newParcels = 1;
        }

        setCellSize(sheet, oldParcels, parcelsAffectedByDPR);

        for (int i = 0; i <= (newParcels* aParcelOrADprNeedsTwoRows + rowsBesideParcelRows); i++){
            Row row =sheet.createRow(i);

            if (i==0) {

                stylingFirstParcelRow(row, oldParcels, excelTemplate);

            } else if (i==1) {

                stylingSecondParcelRow(row, oldParcels, excelTemplate);

            } else if (i==2) {

                stylingThirdParcelRow(row, oldParcels, excelTemplate);


            } else if (i==newParcels* aParcelOrADprNeedsTwoRows +rowsBesideParcelRows) {

                stylingLastParcelRow(row, oldParcels, excelTemplate);


            } else {

                stylingEveryOtherParcelRow(row, oldParcels, newParcels, i, excelTemplate);

            }
        }

        try {
            FileOutputStream out = new FileOutputStream(new File(filePath));
            excelTemplate.write(out);
            out.close();

        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Could not create parcel table " + e.getMessage());
            throw new RuntimeException(e);

        }

    }

    /**
     * adds merged region to excel sheet
     * @param sheet         excel sheet
     * @param oldParcels    amount of old parcels
     */
    private void mergeCellsToRegions(XSSFSheet sheet, int oldParcels) {
        if (oldParcels > 1){
            sheet.addMergedRegion(new CellRangeAddress(0,0,1, oldParcels));
            sheet.addMergedRegion(new CellRangeAddress(1,1,1, oldParcels));
        }
    }

    /**
     * sets cell size for excel sheet
     * @param sheet                 excel sheet
     * @param oldParcels            amount of old parcels
     * @param parcelsAffectedByDPR  amount of parcels in dpr table
     */
    private void setCellSize(XSSFSheet sheet, int oldParcels, int parcelsAffectedByDPR){

        setDefaultCellSize(sheet);
        setColumnWidth(sheet, oldParcels, parcelsAffectedByDPR);
    }

    /**
     * sets default cell size
     * @param sheet     excel sheet
     */
    private void setDefaultCellSize(XSSFSheet sheet){
        sheet.setDefaultRowHeight(defaultRowHeight);
        sheet.setDefaultColumnWidth(defaultColumnWidth);

    }

    /**
     * sets column height
     * @param sheet                 excel sheet
     * @param oldParcels            amount of old parcels in parcel table
     * @param parcelsAffectedByDPR  amount of parcels in dpr table
     */
    private void setColumnWidth(XSSFSheet sheet, int oldParcels, int parcelsAffectedByDPR) {
        sheet.setColumnWidth(0, widthOfFirstColumn);
        if (oldParcels >= parcelsAffectedByDPR) {
            sheet.setColumnWidth(oldParcels + additionConstantToGetIndexOfLastColumnOfParcelTable,
                    widthOfLastParcelTableColumnAndTheFollowingColumn);
            sheet.setColumnWidth(oldParcels + additionConstantToGetIndexOfNewAreaColumnOfDPRTable,
                    widthOfLastParcelTableColumnAndTheFollowingColumn);
        }

    }

    /**
     * styles the first row in parcel table
     * @param row               first row in parcel table
     * @param oldParcels        amount of old parcels
     * @param excelTemplate     excel workbook
     */
    private void  stylingFirstParcelRow(Row row, int oldParcels, XSSFWorkbook excelTemplate) {

        for (int c = 1; c <= oldParcels; c++){
            cell = row.createCell(c);
            cell.setCellValue(oldParcelString);

            color = lightGrayString;
            border_bottom = noStyling;
            border_top = thick;
            border_left = noStyling;
            border_right = noStyling;

            if (c==1) {
                if (oldParcels == 1) {
                    border_left = thick;
                    border_right = thick;
                } else if (oldParcels > 1) {
                    border_left = thick;
                }
            } else if (c==oldParcels) {
                border_right = thick;
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    noIndentValue, excelTemplate);
            cell.setCellStyle(newStyle);
        }

        row.setHeight(rowHeightOfTableHeaderAndFooter);

    }

    /**
     * creates a cell specific styling
     * @param color             background color of cell
     * @param border_bottom     cell border styling (bottom)
     * @param border_top        cell border styling (top)
     * @param border_left       cell border styling (left)
     * @param border_right      cell border styling (right)
     * @param indent            text indent
     * @param excelTemplate     excel workbook
     * @return                  cell style
     */
    private XSSFCellStyle getStyleForCell(String color, String border_bottom, String border_top, String border_left,
                                          String border_right, int indent, XSSFWorkbook excelTemplate ) {

        XSSFCellStyle style = excelTemplate.createCellStyle();

        XSSFColor lightGray = new XSSFColor(new java.awt.Color(217, 217,217));

        XSSFFont font = excelTemplate.createFont();
        font.setFontHeightInPoints(fontHeight);
        font.setFontName(fontName);
        font.setItalic(false);

        style.setFont(font);


        if (color.equals(lightGrayString)){
            style.setFillForegroundColor(lightGray);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setWrapText(true);
        } else {
            style.setVerticalAlignment(VerticalAlignment.BOTTOM);
            style.setAlignment(HorizontalAlignment.RIGHT);
            style.setIndention((short) indent);
        }

        switch (border_bottom) {
            case thick:
                style.setBorderBottom(BorderStyle.THICK);
                break;
            case thin:
                style.setBorderBottom(BorderStyle.THIN);
                break;
            case noStyling:
                style.setBorderBottom(BorderStyle.NONE);
        }

        switch (border_top) {
            case thick:
                style.setBorderTop(BorderStyle.THICK);
                break;
            case thin:
                style.setBorderTop(BorderStyle.THIN);
                break;
            case noStyling:
                style.setBorderTop(BorderStyle.NONE);
        }

        switch (border_left) {
            case thick:
                style.setBorderLeft(BorderStyle.THICK);
                break;
            case thin:
                style.setBorderLeft(BorderStyle.THIN);
                break;
        }


        switch (border_right) {
            case thick:
                style.setBorderRight(BorderStyle.THICK);
                break;
            case thin:
                style.setBorderRight(BorderStyle.THIN);
                break;
        }

        return style;

    }

    /**
     * styles the second row in parcel table
     * @param row               second row
     * @param oldParcels        amount of old parcels
     * @param excelTemplate     excel workbook
     */
    private void stylingSecondParcelRow(Row row, int oldParcels, XSSFWorkbook excelTemplate){

        int totalSize = oldParcels + additionConstantToGetIndexOfLastColumnOfParcelTable;

        for (int c = 0; c <= totalSize; c++){
            cell = row.createCell(c);
            color = lightGrayString;
            border_bottom = thin;
            border_top = thin;
            border_left = thick;
            border_right = thick;

            if (c==0){
                border_top = thick;
                cell.setCellValue(newParcelString);
            } else if (c==1) {
                if (oldParcels > 1) {
                    border_right = thin;
                }
                cell.setCellValue(parcelNumberString);
            } else if (c==oldParcels){
                border_left = thin;
            } else if (c<oldParcels) {
                border_left = thin;
                border_right = thin;
            } else if (c==totalSize){
                border_top = thick;
                cell.setCellValue(newAreaString);
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    noIndentValue, excelTemplate);
            cell.setCellStyle(newStyle);
        }


        row.setHeight(rowHeightOfTableHeaderAndFooter);
    }

    /**
     * styles the third row in parcel table
     * @param row               third row
     * @param oldParcels        amount of old parcels
     * @param excelTemplate     excel workbook
     */
    private void stylingThirdParcelRow (Row row, int oldParcels, XSSFWorkbook excelTemplate){

        int totalSize = oldParcels + additionConstantToGetIndexOfLastColumnOfParcelTable;

        for (int c = 0; c <= totalSize; c++){
            cell = row.createCell(c);

            color = noStyling;
            border_bottom = thick;
            border_top = thin;
            border_left = thick;
            border_right = thick;
            indent = indentValue;


            if (c==0){
                color = lightGrayString;
                indent = noIndentValue;
                cell.setCellValue(parcelNumberString);
            } else if (c==1) {
                cell.setCellType(CellType.STRING);
                if (oldParcels > 1) {
                    border_right = thin;
                }
            } else if (c==oldParcels){
                cell.setCellType(CellType.STRING);
                border_left = thin;
            } else if (c<oldParcels) {
                cell.setCellType(CellType.STRING);
                border_left = thin;
                border_right = thin;
            } else if (c==totalSize){
                color = lightGrayString;
                indent = noIndentValue;
                cell.setCellValue(squareMeterString);
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            cell.setCellStyle(newStyle);
        }

        row.setHeight(rowHeightOfTableHeaderAndFooter);
    }

    /**
     * styles the last row of parcel table
     * @param row               last row
     * @param oldParcels        amount of old parcels
     * @param excelTemplate     excel workbook
     */
    private void stylingLastParcelRow(Row row, int oldParcels, XSSFWorkbook excelTemplate){

        int totalSize = oldParcels + additionConstantToGetIndexOfLastColumnOfParcelTable;

        for (int c = 0; c <= totalSize; c++) {
            cell = row.createCell(c);

            color = noStyling;
            border_bottom = thick;
            border_top = thick;
            border_left = thick;
            border_right = thick;
            indent = indentValue;

            if (c == 0) {
                color = lightGrayString;
                indent = noIndentValue;
                cell.setCellValue(oldAreaString);
            } else if (c == 1) {
                if (oldParcels > 1) {
                    border_right = thin;
                }
            } else if (c == oldParcels) {
                border_left = thin;
            } else if (c < oldParcels) {
                border_left = thin;
                border_right = thin;
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            cell.setCellStyle(newStyle);
        }

        row.setHeight(rowHeightOfTableHeaderAndFooter);
    }

    /**
     * styles every row between the third and the last row of parcel table
     * @param row               row
     * @param oldParcels        amount of old parcels
     * @param newParcels        amount of new parcels
     * @param i                 row index
     * @param excelTemplate     excel workbook
     */
    private void stylingEveryOtherParcelRow(Row row, int oldParcels, int newParcels, int i, XSSFWorkbook excelTemplate){

        if (i % aParcelOrADprNeedsTwoRows == 0){
            border_bottom = thin;
            border_top = noStyling;
        } else {
            border_bottom = noStyling;
            border_top = thin;
        }


        int totalSize = oldParcels + additionConstantToGetIndexOfLastColumnOfParcelTable;

        for (int c = 0; c <= totalSize; c++) {
            cell = row.createCell(c);

            color = noStyling;
            border_left = thick;
            border_right = thick;
            indent = indentValue;

            if (c == 0) {
                cell.setCellType(CellType.STRING);
                if (i==newParcels * aParcelOrADprNeedsTwoRows + rowsBesideParcelRows - 1){
                    border_bottom = thick;
                    border_top = noStyling;
                    indent = noIndentValue;
                    cell.setCellValue(roundingDifferenceString);
                }
            } else if (c == 1) {
                if (oldParcels > 1) {
                    border_right = thin;
                }
            } else if (c == oldParcels) {
                border_left = thin;
            } else if (c < oldParcels) {
                border_left = thin;
                border_right = thin;
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            cell.setCellStyle(newStyle);
        }
    }


    /**
     * creates the dpr table
     * @param excelTemplate     excel workbook
     * @param filePath          path, where excel file should be saved to
     * @param parcels           amount of parcels in dpr table
     * @param dpr               amount of dprs in dpr table
     * @param newParcels        amount of new parcels in parcel table
     * @param oldParcels        amount of old parcels in parcel table
     */
    @Override
    public void createDPRTable(XSSFWorkbook excelTemplate, String filePath, int parcels, int dpr,
                                       int newParcels, int oldParcels) {

        if (newParcels == 0 ){
            newParcels = 1;
        } else if (oldParcels == 0){
            oldParcels = 1;
        }

        if (dpr==0){
            dpr = 1;
            parcels = 1;
        } else if (parcels == 0){
            dpr = 1;
            parcels = 1;
        }

        int rowStartIndex = (9 + aParcelOrADprNeedsTwoRows * newParcels - 1);

        XSSFSheet sheet = excelTemplate.getSheet(xlsxSheetName);

        setColumnWidth(parcels, oldParcels, sheet);

        addMergedRegionsDPR(rowStartIndex, parcels, sheet);


        for (int i = rowStartIndex; i < (rowStartIndex + 3 + aParcelOrADprNeedsTwoRows * dpr); i++) {
            Row row = sheet.createRow(i);
            if (i == rowStartIndex) {

                stylingFirstDPRRow(row, parcels, excelTemplate);

            } else if (i == rowStartIndex + 1) {

                stylingSecondDPRRow(row, parcels, oldParcels, excelTemplate);

            } else if (i == rowStartIndex + 2) {

                stylingRowWithDPRNumber(row, parcels, excelTemplate);

            }  else {

                stylingEveryOtherDPRRow(row, i, rowStartIndex, parcels, dpr, excelTemplate);

            }
        }


        try {
            FileOutputStream out = new FileOutputStream(new File(filePath));
            excelTemplate.write(out);
            out.close();
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Could not create dpr table " + e.getMessage());
            throw new RuntimeException(e);
        }

    }

    /**
     * sets column width for the two last columns
     * @param parcels       amount of parcels in dpr table
     * @param oldParcels    amount of old parcels in parcel table
     * @param sheet         excel sheet
     */
    private void setColumnWidth(int parcels, int oldParcels, XSSFSheet sheet){
        if (parcels > oldParcels) {
            sheet.setColumnWidth(parcels + additionConstantToGetIndexOfLastColumnOfParcelTable,
                    widthOfLastParcelTableColumnAndTheFollowingColumn);
            sheet.setColumnWidth(parcels + additionConstantToGetIndexOfNewAreaColumnOfDPRTable,
                    widthOfLastParcelTableColumnAndTheFollowingColumn);
        }
    }

    /**
     * adds merged regions to excel
     * @param rowStartIndex     start index for row
     * @param parcels           amount of parcels in dpr table
     * @param sheet             excel sheet
     */
    private void addMergedRegionsDPR(int rowStartIndex, int parcels, XSSFSheet sheet){
        if (parcels>1){
            sheet.addMergedRegion(new CellRangeAddress(rowStartIndex, rowStartIndex,1, parcels));
            sheet.addMergedRegion(new CellRangeAddress(rowStartIndex + 1,rowStartIndex + 1,
                    1, parcels));
        }

        sheet.addMergedRegion((new CellRangeAddress(rowStartIndex + 1, rowStartIndex + 2,
                parcels + additionConstantToGetIndexOfRoundingDifferenceColumnOfDPRTable,
                parcels + additionConstantToGetIndexOfRoundingDifferenceColumnOfDPRTable)));
    }

    /**
     * styles the first row in dpr table
     * @param row               first row
     * @param parcels           amount of parcels in dpr table
     * @param excelTemplate     excel workbook
     */
    private void stylingFirstDPRRow(Row row, int parcels, XSSFWorkbook excelTemplate){

        for (int c = 1; c <= parcels; c++) {
            cell = row.createCell(c);
            cell.setCellValue(parcelString);

            color = lightGrayString;
            border_bottom = noStyling;
            border_top = thick;
            border_left = noStyling;
            border_right = noStyling;

            if (c == 1) {
                if (parcels == 1) {
                    border_left = thick;
                    border_right = thick;
                } else if (parcels > 1) {
                    border_left=thick;
                }
            } else if (c == parcels) {
                border_right = thick;
            }

            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    noIndentValue, excelTemplate);
            cell.setCellStyle(newStyle);

        }
        row.setHeight(rowHeightOfTableHeaderAndFooter);
    }

    /**
     * styles the second row in dpr table
     * @param row               second row
     * @param parcels           amount of parcels in dpr table
     * @param oldParcels        amount of old parcels in parcel table
     * @param excelTemplate     excel workbook
     */
    private void stylingSecondDPRRow (Row row, int parcels, int oldParcels, XSSFWorkbook excelTemplate){

        int totalSize = parcels + additionConstantToGetIndexOfNewAreaColumnOfDPRTable;

        for (int c = 0; c <= totalSize; c++){
            cell = row.createCell(c);

            color = lightGrayString;
            border_bottom = thin;
            border_top = thin;
            border_left = thin;
            border_right = thin;

            if (c==0){
                border_top = thick;
                border_left = thick;
                border_right = thick;
                cell.setCellValue(dprString);

            } else if (c==1) {
                border_left = thick;
                cell.setCellValue(parcelNumberString);
            } else if (c==parcels + additionConstantToGetIndexOfRoundingDifferenceColumnOfDPRTable){
                border_bottom = noStyling;
                border_top = thick;
                border_right = thick;
                if (parcels >= oldParcels) {
                    cell.setCellValue(splittedRoundingDifferenceString);
                } else {
                    cell.setCellValue(roundingDifferenceString);
                }
            }else if (c==totalSize){
                border_top = thick;
                border_left = thick;
                border_right = thick;
                cell.setCellValue(dprAreaString);
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    noIndentValue, excelTemplate);
            if (c==0){
                newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            }
            cell.setCellStyle(newStyle);
        }

        row.setHeight( rowHeightOfTableHeaderAndFooter);
    }

    /**
     * styles every row where a dpr number is written
     * @param row               row with dpr number
     * @param parcels           amount of parcels in dpr table
     * @param excelTemplate     excel workbook
     */
    private void stylingRowWithDPRNumber(Row row, int parcels, XSSFWorkbook excelTemplate) {

        int totalSize = parcels + additionConstantToGetIndexOfNewAreaColumnOfDPRTable;
        for (int c = 0; c <= totalSize; c++) {
            cell = row.createCell(c);

            color = lightGrayString;
            border_bottom = thick;
            border_top = thin;
            border_left = thin;
            border_right = thick;
            indent = noIndentValue;


            if (c == 0) {
                border_left = thick;
                cell.setCellValue(parcelNumberString);

            } else if (c == 1) {
                color = "";
                border_left = thick;
                border_right = thin;
                indent = 2;
            } else if (c <= parcels && parcels != 1) {
                color = noStyling;
                border_right = thin;
                indent = 2;
            } else if (c == parcels + additionConstantToGetIndexOfRoundingDifferenceColumnOfDPRTable) {
                border_top = noStyling;
            } else if (c == totalSize) {
                border_left = thick;
                cell.setCellValue(squareMeterString);
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indent, excelTemplate);
            if (c==0){
                newStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            }
            cell.setCellStyle(newStyle);
        }

        row.setHeight(rowHeightOfTableHeaderAndFooter);
    }

    /**
     * styles all other rows that are not first row or second row or a row with dpr
     * @param row               row
     * @param i                 row index
     * @param rowStartIndex     start index of first row in dpr table
     * @param parcels           amount of parcels in dpr table
     * @param dpr               amount of dprs
     * @param excelTemplate     excel workbook
     */
    private void stylingEveryOtherDPRRow(Row row, int i, int rowStartIndex, int parcels, int dpr,
                                         XSSFWorkbook excelTemplate ) {

        int totalSize = parcels + additionConstantToGetIndexOfNewAreaColumnOfDPRTable;


        if ((i-rowStartIndex) % aParcelOrADprNeedsTwoRows == 0){
            border_bottom = thin;
            border_top = noStyling;
        } else {
            border_bottom = noStyling;
            border_top = thin;
        }
        for (int c = 0; c <= totalSize; c++) {
            cell = row.createCell(c);

            color = noStyling;
            border_left = thick;
            border_right = thick;

            if (c == 0) {
                if (i==dpr * aParcelOrADprNeedsTwoRows + rowStartIndex + 2){
                    border_bottom = thick;
                    border_top = noStyling;
                    cell.setCellType(CellType.STRING);
                } else {
                    if (border_bottom.equals(thin)){
                        cell.setCellType(CellType.STRING);
                    }
                }
            } else if (c == 1) {
                if (i==dpr * aParcelOrADprNeedsTwoRows + rowStartIndex + 2){
                    border_bottom = thick;
                    border_top = noStyling;
                    border_right = thin;
                } else {
                    border_right = thin;
                }
            } else if (c <= parcels && parcels != 1) {
                if (i==dpr * aParcelOrADprNeedsTwoRows + rowStartIndex + 2){
                    border_bottom = thick;
                    border_top = noStyling;
                    border_left = thin;
                    border_right = thin;
                } else {
                    border_left = thin;
                    border_right = thin;
                }
            } else if (c == parcels + additionConstantToGetIndexOfRoundingDifferenceColumnOfDPRTable) {
                if (i==dpr * aParcelOrADprNeedsTwoRows + rowStartIndex + 2){
                    border_bottom = thick;
                    border_top = noStyling;
                    border_left = thin;
                } else {
                    border_left = thin;
                }
            } else if (c == totalSize) {
                if (i==dpr * aParcelOrADprNeedsTwoRows + rowStartIndex + 2){

                    border_bottom = thick;
                    border_top = noStyling;
                }
            }
            XSSFCellStyle newStyle = getStyleForCell(color, border_bottom, border_top, border_left, border_right,
                    indentValue, excelTemplate);
            cell.setCellStyle(newStyle);
        }
    }

}
