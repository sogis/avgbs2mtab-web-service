package ch.so.agi.avgbs2mtab.writeexcel;


import ch.so.agi.avgbs2mtab.writeexcel.XLSXTemplate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.util.ArrayList;
import java.util.List;


public class XLSXTemplateTest {

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();

    @Test
    public void workbookCreatedOnWritablePath() throws Exception {

        File excelFile = folder.newFile("test.xlsx");

        String filePath = excelFile.getAbsolutePath();
        XLSXTemplate xlsxTemplate = new XLSXTemplate();

        xlsxTemplate.createWorkbook(filePath);
    }


    @Test
    public void correctStyledCellsInExcel() throws Exception {
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();

        List<Integer> oldParcels = generateOldParcels();
        List<Integer> newParcels = generateNewParcels();
        List<Integer> parcels = generateParcels();
        List<Integer> dpr = generateDPR();

        try {
            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
            xlsxTemplate.createParcelTable(newWorkbook, filePath, newParcels.size(), oldParcels.size(), parcels.size());
            xlsxTemplate.createDPRTable(newWorkbook, filePath, parcels.size(), dpr.size(),
                    newParcels.size(), oldParcels.size());

            Assert.assertTrue(checkStylingOfExcelCells(newWorkbook));

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    @Test
    public void correctStyledCellsInExcelWithoutAnyParcelsAndDPRs() throws Exception {
        File excelFile = folder.newFile("test.xlsx");
        String filePath = excelFile.getAbsolutePath();

        XLSXTemplate xlsxTemplate = new XLSXTemplate();

        int oldParcels = 0;
        int newParcels = 0;
        int parcels = 0;
        int dpr = 0;

        try {
            XSSFWorkbook newWorkbook = xlsxTemplate.createWorkbook(filePath);
            xlsxTemplate.createParcelTable(newWorkbook, filePath, newParcels, oldParcels, parcels);
            xlsxTemplate.createDPRTable(newWorkbook, filePath, parcels, dpr,
                    newParcels, oldParcels);

            Assert.assertTrue(checkStylingOfExcelCellsWithoutAnyParcelsAndDPRs(newWorkbook));

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    private List<Integer> generateOldParcels() {
        List<Integer> oldparcels = new ArrayList<>();
        oldparcels.add(695);
        oldparcels.add(696);
        oldparcels.add(697);
        oldparcels.add(701);
        oldparcels.add(870);
        oldparcels.add(874);

        return oldparcels;
    }

    private List<Integer> generateNewParcels() {
        List<Integer> newParcels = new ArrayList<>();
        newParcels.add(695);
        newParcels.add(696);
        newParcels.add(697);
        newParcels.add(701);
        newParcels.add(870);
        newParcels.add(874);
        newParcels.add(4004);

        return newParcels;
    }

    private List<Integer> generateParcels() {
        List<Integer> parcels = new ArrayList<>();
        parcels.add(2174);
        parcels.add(2175);
        parcels.add(2176);
        parcels.add(2174);
        parcels.add(2175);
        parcels.add(2176);
        parcels.add(2174);
        parcels.add(2175);
        parcels.add(2176);

        return parcels;
    }

    private List<Integer> generateDPR() {
        List<Integer> dpr = new ArrayList<>();
        dpr.add(40053);
        dpr.add(15828);

        return dpr;
    }

    private boolean checkStylingOfExcelCells(XSSFWorkbook workbook) {

        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allCellsAreCorrectlyStyled = true;

        XSSFColor lightGray = new XSSFColor(new java.awt.Color(217, 217, 217));

        for (int i = 0; i < 20; i++) {
            XSSFRow row = xlsxSheet.getRow(i);
            if (i == 0) {
                for (int c = 1; c <= 6; c++) {
                    XSSFCell cell = row.getCell(c);
                    if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray)) {
                        allCellsAreCorrectlyStyled = false;
                    }
                }
            } else if (i == 1) {
                for (int c = 0; c <= 6 + 1; c++) {
                    XSSFCell cell = row.getCell(c);
                    if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray)) {
                        allCellsAreCorrectlyStyled = false;
                    }
                }

            } else if (i == 2) {
                XSSFCell cell = row.getCell(0);
                if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray) ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THICK) ||
                        cell.getCellStyle().getIndention() != 0 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.CENTER) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.CENTER) ||
                        !cell.getStringCellValue().equals("Grundstück-Nr.")) {
                    allCellsAreCorrectlyStyled = false;
                }

                cell = row.getCell(1);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THIN) ||
                        cell.getCellStyle().getIndention() != 2 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM)) {
                    allCellsAreCorrectlyStyled = false;
                }

            } else if (i == 6) {
                XSSFCell cell = row.getCell(3);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.NONE) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THIN) ||
                        cell.getCellStyle().getIndention() != 2 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM)) {
                    allCellsAreCorrectlyStyled = false;
                }


            } else if (i == 18) {
                XSSFCell cell = row.getCell(0);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.NONE) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THICK) ||
                        cell.getCellStyle().getIndention() != 0 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM) ||
                        !cell.getStringCellValue().equals("Rundungsdifferenz")) {
                    allCellsAreCorrectlyStyled = false;
                }

            }
        }

        if (!xlsxSheet.getMergedRegions().contains(new CellRangeAddress(0, 0, 1, 6)) ||
                !xlsxSheet.getMergedRegions().contains(new CellRangeAddress(1, 1, 1, 6))) {
            allCellsAreCorrectlyStyled = false;
        }

        return allCellsAreCorrectlyStyled;

    }

    private boolean checkStylingOfExcelCellsWithoutAnyParcelsAndDPRs(XSSFWorkbook workbook) {

        XSSFSheet xlsxSheet = workbook.getSheet("Mutationstabelle");

        boolean allCellsAreCorrectlyStyled = true;

        XSSFColor lightGray = new XSSFColor(new java.awt.Color(217, 217, 217));

        for (int i = 0; i <= 14; i++) {
            XSSFRow row = xlsxSheet.getRow(i);
            if (i == 0) {
                int c = 1;
                XSSFCell cell = row.getCell(c);
                if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray)) {
                    allCellsAreCorrectlyStyled = false;
                }
            } else if (i == 1) {
                for (int c = 0; c <= 1 + 1; c++) {
                    XSSFCell cell = row.getCell(c);
                    if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray)) {
                        allCellsAreCorrectlyStyled = false;
                    }
                }

            } else if (i == 2) {
                XSSFCell cell = row.getCell(0);
                if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray) ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THICK) ||
                        cell.getCellStyle().getIndention() != 0 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.CENTER) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.CENTER)) {
                    allCellsAreCorrectlyStyled = false;
                }

                cell = row.getCell(1);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THICK) ||
                        cell.getCellStyle().getIndention() != 2 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM)) {
                    allCellsAreCorrectlyStyled = false;
                }

            } else if (i == 4) {
                XSSFCell cell = row.getCell(1);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THIN) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.NONE) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THICK) ||
                        cell.getCellStyle().getIndention() != 2 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM)) {
                    allCellsAreCorrectlyStyled = false;
                }


            } else if (i == 6) {
                XSSFCell cell = row.getCell(0);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.NONE) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THICK) ||
                        cell.getCellStyle().getIndention() != 0 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM)) {
                    allCellsAreCorrectlyStyled = false;
                }

            } else if (i == 11) {
                for (int c = 0; c <= 1 + 2; c++) {
                    XSSFCell cell = row.getCell(c);
                    if (!cell.getCellStyle().getFillForegroundXSSFColor().equals(lightGray)) {
                        allCellsAreCorrectlyStyled = false;
                    }
                }
            } else if (i == 14) {
                XSSFCell cell = row.getCell(1);
                if (cell.getCellStyle().getFillForegroundXSSFColor() != null ||
                        !cell.getCellStyle().getBorderBottomEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderTopEnum().equals(BorderStyle.NONE) ||
                        !cell.getCellStyle().getBorderLeftEnum().equals(BorderStyle.THICK) ||
                        !cell.getCellStyle().getBorderRightEnum().equals(BorderStyle.THIN) ||
                        cell.getCellStyle().getIndention() != 2 ||
                        !cell.getCellStyle().getFont().getFontName().equals("Arial") ||
                        cell.getCellStyle().getFont().getFontHeightInPoints() != 11 ||
                        !cell.getCellStyle().getAlignmentEnum().equals(HorizontalAlignment.RIGHT) ||
                        !cell.getCellStyle().getVerticalAlignmentEnum().equals(VerticalAlignment.BOTTOM)) {
                    allCellsAreCorrectlyStyled = false;
                }

            }
        }

        if (!xlsxSheet.getMergedRegions().contains(new CellRangeAddress(11, 12, 2, 2))) {
            allCellsAreCorrectlyStyled = false;
        }

        if (!xlsxSheet.getRow(0).getCell(1).getStringCellValue().equals("Alte Liegenschaften") ||
                !xlsxSheet.getRow(1).getCell(0).getStringCellValue().equals("Neue Liegenschaften") ||
                !xlsxSheet.getRow(1).getCell(1).getStringCellValue().equals("Grundstück-Nr.") ||
                !xlsxSheet.getRow(1).getCell(2).getStringCellValue().equals("Neue Fläche") ||
                !xlsxSheet.getRow(2).getCell(0).getStringCellValue().equals("Grundstück-Nr.") ||
                !xlsxSheet.getRow(2).getCell(2).getStringCellValue().equals("[m2]") ||
                !xlsxSheet.getRow(6).getCell(0).getStringCellValue().equals("Rundungsdifferenz") ||
                !xlsxSheet.getRow(7).getCell(0).getStringCellValue().equals("Alte Fläche [m2]") ||
                !xlsxSheet.getRow(10).getCell(1).getStringCellValue().equals("Liegenschaften") ||
                !xlsxSheet.getRow(11).getCell(0).getStringCellValue().equals("Selbst. Recht") ||
                !xlsxSheet.getRow(11).getCell(1).getStringCellValue().equals("Grundstück-Nr.") ||
                !xlsxSheet.getRow(11).getCell(2).getStringCellValue().equals("Rundungs-differenz") ||
                !xlsxSheet.getRow(11).getCell(3).getStringCellValue().equals("Selbst. Recht Fläche") ||
                !xlsxSheet.getRow(12).getCell(0).getStringCellValue().equals("Grundstück-Nr.") ||
                !xlsxSheet.getRow(12).getCell(3).getStringCellValue().equals("[m2]")) {
            allCellsAreCorrectlyStyled = false;
        }


        return allCellsAreCorrectlyStyled;

    }
}

