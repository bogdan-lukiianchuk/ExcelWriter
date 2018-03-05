package com.bogluk.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class Main {
    private static Map<String, CellStyle> styles = new HashMap<>();

    private  static Map<String, Object[]> data = new HashMap<String, Object[]>();
    static {
        data.put("7", new Object[]{7d, "Sonya", "75K", "SALES", "Rupert"});
        data.put("8", new Object[]{8d, "Kris", "85K", "SALES", "Rupert"});
        data.put("9", new Object[]{9d, "Dave", "90K", "SALES", "Rupert"});
    }

    public static void main(String[] args) {
//        Workbook book = new HSSFWorkbook(); // xls
        Workbook book = new XSSFWorkbook(); // xlsx
        Sheet mySheet = book.createSheet("Orders");
        styles = createStyles(book);
        addHeader(mySheet, "Club", styles.get("header"));
        mySheet.createRow(mySheet.getLastRowNum() + 1);

        String[] columnHeaders = new String[] {"Number", "Name", "Salary", "Department", "Manager"};
        final Row columnHeaderRow = mySheet.createRow(mySheet.getLastRowNum() + 1);
        int column = 0;
        for (String s: columnHeaders) {
            final Cell cell = columnHeaderRow.createCell(column++);
            cell.setCellValue(s);
            cell.setCellStyle(styles.get("column header"));
        }

        Set<String> newRows = data.keySet();
        // get the last row number to append new data
        int rownum = mySheet.getLastRowNum() + 1;
        for (String key : newRows) { // Creating a new Row in existing XLSX sheet Row
            final Row row = mySheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellCount = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellCount++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Boolean) {
                    cell.setCellValue((Boolean) obj);
                } else if (obj instanceof Date) {
                    cell.setCellValue((Date) obj);
                } else if (obj instanceof Double) {
                    cell.setCellValue((Double) obj);
                }
            }
        }
        final Row row = mySheet.createRow(rownum++);
        final Cell cell = row.createCell(0);
        cell.setCellFormula("SUM(A4:A6)");
        // open an OutputStream to save written data into XLSX file FileOutputStream os = new FileOutputStream(myFile);
        File newFile = new File("F://temp//employee1.xlsx");
        try (OutputStream os = new FileOutputStream(newFile)){
            book.write(os);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Writing on XLSX file Finished ...");
    }

    private static void addHeader(Sheet sheet, String header, CellStyle style) {
        int last = sheet.getLastRowNum();
        final Workbook workbook = sheet.getWorkbook();
        final Row row = sheet.createRow(last);
        final Cell cell = row.createCell(0);
        cell.setCellValue(header);
        cell.setCellStyle(style);
    }

    private static Map<String, CellStyle> createStyles(Workbook workbook) {
        Map<String, CellStyle> styles = new HashMap<>();
        final CellStyle headerCellStyle = workbook.createCellStyle();
        final Font headerCellFont = workbook.createFont();
        headerCellFont.setBold(true);
        headerCellFont.setFontHeightInPoints((short) 16);
        headerCellStyle.setFont(headerCellFont);
        styles.put("header", headerCellStyle);

        final CellStyle columnHeaderCellStyle = workbook.createCellStyle();
        final Font columnHeaderCellFont = workbook.createFont();
        columnHeaderCellFont.setBold(true);
        columnHeaderCellStyle.setFont(columnHeaderCellFont);
        columnHeaderCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        columnHeaderCellStyle.setBorderTop(BorderStyle.MEDIUM);
        columnHeaderCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        columnHeaderCellStyle.setBorderRight(BorderStyle.MEDIUM);
        styles.put("column header", columnHeaderCellStyle);

        return styles;
    }

}
