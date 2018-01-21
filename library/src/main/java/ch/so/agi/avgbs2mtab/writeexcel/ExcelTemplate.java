package ch.so.agi.avgbs2mtab.writeexcel;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public interface ExcelTemplate {

    public XSSFWorkbook createWorkbook(String filePath);

    public void createParcelTable(XSSFWorkbook excelTemplate, String filePath, int newParcels, int oldParcels,
                                          int parcelsAffectedByDPR);

    public void createDPRTable(XSSFWorkbook excelTemplate, String filePath, int parcels, int dpr,
                                       int newParcels, int oldParcels);
}
