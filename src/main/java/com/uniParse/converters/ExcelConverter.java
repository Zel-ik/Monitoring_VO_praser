package com.uniParse.converters;

import com.uniParse.Entity.University;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.jsoup.internal.StringUtil;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelConverter {
    public void createWorkbook(String excelFileName, List<University> universities) {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet;
        CellStyle style = getStyle(wb);
        try (OutputStream fileOut = new FileOutputStream(excelFileName + ".xls")) {
            for (int i = 0; i < universities.size(); i++) {

                sheet = wb.createSheet(String.valueOf(i));
                Row row = sheet.createRow(1);
                row.setHeight((short) 600);

                Cell cell = row.createCell(1);
                cell.setCellValue(universities.get(i).getName());


                cell.setCellStyle(style);
                sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 6));
                setBordersToMergedCells(sheet);
            }
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void writeTablesIntoExcel(String excelFileName,
                                     String universityNumber,
                                     Elements resultTable,
                                     Elements tablesWithAttributeNapde,
                                     Elements tableKontByOtr,
                                     Elements tableAnalisReq,
                                     Elements tableAnalisDop) {
        try {
            if (resultTable != null) {
                FileInputStream file = new FileInputStream(excelFileName + ".xls");
                Workbook wb = new HSSFWorkbook(file);
                Sheet sheet = wb.getSheet(universityNumber);
                sheet.setDefaultColumnWidth(35);
                int rowNumber = 3;

                resultTable.select("table > tbody > thead");
                rowNumber = writeRows(resultTable, sheet, rowNumber);

                for (Element element : tablesWithAttributeNapde) {
                    rowNumber = writeRows(element, sheet, rowNumber);
                }

                rowNumber = writeRows(tableKontByOtr, sheet, rowNumber);
                rowNumber = writeRows(tableAnalisReq, sheet, rowNumber);
                rowNumber = writeRows(tableAnalisDop, sheet, rowNumber);

                FileOutputStream fileOutputStream = new FileOutputStream(excelFileName + ".xls");
                wb.write(fileOutputStream);
                fileOutputStream.flush();
                wb.close();
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }

    private Integer writeRows(Elements table, Sheet sheet, int rowNumber) {

        for (Element row : table.select("tr")) {
            Row exlrow = sheet.createRow(rowNumber++);
            exlrow.setHeight((short) -1);
            writeColumns(rowNumber, row, exlrow, sheet);
        }

        return rowNumber + 5;
    }

    private Integer writeRows(Element table, Sheet sheet, int rowNumber) {

        for (Element row : table.select("tr")) {
            Row exlrow = sheet.createRow(rowNumber++);
            exlrow.setHeight((short) -1);
            writeColumns(rowNumber, row, exlrow, sheet);
        }
        return rowNumber + 5;
    }

    private void writeColumns(int rowNumber, Element row, Row exlrow, Sheet sheet) {
        int cellnum = 1;
        for (Element tds : row.select("td")) {
            Cell cell = exlrow.createCell(cellnum++);

            if (StringUtil.isNumeric(tds.text())) cell.setCellValue(Integer.parseInt(tds.text()));

            else if (StringUtil.isNumeric(tds.text()
                    .replace(",", "")
                    .replace(" ", ""))) {
                cell.setCellValue(Double.parseDouble(tds.text().replace(",", ".").replace(" ", "")));
                CellUtil.setCellStyleProperties(cell, getStyleMap());

            } else if (StringUtil.isNumeric(tds.text()
                    .replace(",", "")
                    .replace(" ", "")
                    .replace("%", ""))) {
                cell.setCellValue(Double.parseDouble(tds.text().replace(",", ".").replace(" ", "").replace("%", "")));
            } else cell.setCellValue(tds.text());


            if (tds.hasAttr("colspan")) sheet.addMergedRegion(new CellRangeAddress(rowNumber - 1, rowNumber - 1, 1, 4));

            CellUtil.setCellStyleProperties(cell, getStyleMap());
        }
    }

    private Map<String, Object> getStyleMap(){
        Map<String, Object> properties = new HashMap<String, Object>();
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.THIN);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.THIN);
        properties.put(CellUtil.BORDER_TOP, BorderStyle.THIN);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.THIN);
        properties.put(CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
        properties.put(CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
        properties.put(CellUtil.WRAP_TEXT, true);
        return properties;
    }

    private CellStyle getStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        return style;
    }

    private void setBordersToMergedCells(Sheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
        }
    }

}
