package ch.so.agi.avgbs2mtab.writeexcel;

import ch.so.agi.avgbs2mtab.mutdat.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * The Class XlsxWriter generates the excel template and fills it with the data from the xml
 */
public class XlsxWriter {

    private DataExtractionDPR dataExtractionDPR;
    private DataExtractionParcel dataExtractionParcel;
    private MetadataOfDPRMutation metadataOfDPRMutation;
    private MetadataOfParcelMutation metadataOfParcelMutation;
    private static final Logger LOGGER = Logger.getLogger( XlsxWriter.class.getName());

    public XlsxWriter(DataExtractionParcel dataExtractionParcel, DataExtractionDPR dataExtractionDPR,
                      MetadataOfParcelMutation metadataOfParcelMutation, MetadataOfDPRMutation metadataOfDPRMutation) {
        this.dataExtractionDPR = dataExtractionDPR;
        this.dataExtractionParcel = dataExtractionParcel;
        this.metadataOfDPRMutation = metadataOfDPRMutation;
        this.metadataOfParcelMutation = metadataOfParcelMutation;

    }

    /**
     * Creates an excel file with all values from xml
     * @param filePath  path, where file should be saved to
     */
    public void writeXlsx(String filePath) {

        XLSXTemplate xlsxTemplate = new XLSXTemplate();
        ExcelData excelData = new ExcelData();

        LOGGER.log(Level.CONFIG, "Start generating and writing excel");
        LOGGER.log(Level.FINER, "Start generating excel-template");
        XSSFWorkbook workbook = xlsxTemplate.createExcelTemplate(filePath, metadataOfParcelMutation, metadataOfDPRMutation);

        LOGGER.log(Level.FINER, "finished generating excel-template; start writing parcels into excel-template");
        excelData.fillValuesIntoParcelTable(filePath, workbook, dataExtractionParcel);

        LOGGER.log(Level.FINER,"finished writing parcels into excel-template; start writing dprs into excel-templat");
        excelData.fillValuesIntoDPRTable(filePath, workbook, dataExtractionDPR, metadataOfParcelMutation);

        LOGGER.log(Level.FINER, "finished writing dprs into excel-template");

        LOGGER.log(Level.CONFIG, "avgbs2mtab finished");
    }
}
