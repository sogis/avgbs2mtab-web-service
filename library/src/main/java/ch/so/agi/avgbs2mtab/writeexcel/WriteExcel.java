package ch.so.agi.avgbs2mtab.writeexcel;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.List;


public interface WriteExcel  {

    void writeOldParcelsInTemplate(List<Integer> orderedListOfOldParcelNumbers,
                                          XSSFSheet xlsxSheet);

    void writeNewParcelsInTemplate(List<Integer> orderedListOfNewParcelNumbers,
                                          XSSFSheet xlsxSheet);

    void writeInflowAndOutflowOfOneParcelPair(int oldParcelNumber,
                                                     int newParcelNumber,
                                                     int area,
                                                     XSSFSheet xlsxSheet);

    void writeNewArea(int newParcelNumber,
                             int area,
                             XSSFSheet xlsxSheet);

    void writeRoundingDifference(int oldParcelNumber,
                                        int roundingDifference,
                                        int numberOfNewParcels,
                                        XSSFSheet xlsxSheet);

    void writeSumOfRoundingDifference (int NumberOfNewParcels,
                                              int NumberOfOldParcels,
                                              int roundingDifferenceSum,
                                              XSSFSheet xlsxSheet);


    void writeOldArea(int oldParcelNumber,
                             int oldArea,
                             int numberOfNewParcels,
                             XSSFSheet xlsxSheet);

    void writeAreaSum(HashMap<Integer, Integer> oldAreas,
                             List<Integer> newAreas,
                             int roundingDifference,
                             XSSFSheet xlsxSheet);


    void writeParcelsAffectedByDPRsInTemplate(List<Integer> orderedListOfParcelNumbersAffectedByDPRs,
                                                     int newParcelNumber,
                                                     XSSFSheet xlsxSheet);

    void writeDPRsInTemplate(List<Integer> orderedListOfDPRs,
                                    int newParcelNumber,
                                    XSSFSheet xlsxSheet);

    void writeDPRInflowAndOutflows(int parcelNumberAffectedByDPR,
                                          int dpr,
                                          int area,
                                          int newParcelNumber,
                                          XSSFSheet xlsxSheet);

    void writeNewDPRArea(int dpr,
                                int area,
                                int newParcelNumber,
                                XSSFSheet xlsxSheet);

    void writeDPRRoundingDifference(int dpr,
                                           int roundingDifference,
                                           int newParcelNumber,
                                           XSSFSheet xlsxSheet);



}
