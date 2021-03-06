/*
 * The MIT License
 *
 * Copyright 2021-2022 Takashi Harano
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
package com.libutil.excelreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * Reads the sheet value of the Excel file.
 */
public class ExcelLoader {

  /**
   * Opens an Excel file and returns it as a Workbook object.
   *
   * @param filePath
   *          The Excel file path
   * @return The Workbook object
   * @throws IOException
   *           If an I/O error occurs
   */
  public static XSSFWorkbook openBook(String filePath) throws IOException {
    File file = new File(filePath);
    if (!file.exists()) {
      throw new RuntimeException("Excel file not found: file=" + filePath);
    }
    XSSFWorkbook workbook = null;
    FileInputStream fis = new FileInputStream(filePath);
    try {
      workbook = new XSSFWorkbook(fis);
    } finally {
      fis.close();
    }
    return workbook;
  }

  /**
   * Reads the Excel master management item definition sheet and returns it as a
   * two-dimensional array.
   *
   * @param workbook
   *          The Excel workbook object
   * @param sheetName
   *          The sheet name
   * @return Two-dimensional array of read contents
   */
  public static SheetValues loadSheetValues(XSSFWorkbook workbook, String sheetName) {
    return loadSheetValues(workbook, sheetName, 0, null);
  }

  /**
   * Reads an Excel sheet and returns it as a two-dimensional array.
   *
   * @param workbook
   *          The Excel workbook object
   * @param sheetName
   *          The sheet name
   * @param lastRowNum
   *          Last line to load (1-1048576) / 0 for auto detection.
   * @param lastCol
   *          Last column to load (A-XFD) / null for auto detection.
   * @return Two-dimensional array of read contents
   */
  public static SheetValues loadSheetValues(XSSFWorkbook workbook, String sheetName, int lastRowNum, String lastCol) {
    int lastCellNum = 0;
    if (lastCol != null) {
      lastCellNum = ExcelStringUtil.xlscol(lastCol);
    }
    return loadSheetValues(workbook, sheetName, lastRowNum, lastCellNum);
  }

  /**
   * Reads an Excel sheet and returns it as a two-dimensional array.
   *
   * @param workbook
   *          The Excel workbook object
   * @param sheetName
   *          The sheet name
   * @param lastRowNum
   *          Last line to load (1-1048576) / 0 for auto detection
   * @param lastCellNum
   *          Last column to load (1-16384) / 0 for auto detection
   * @return Two-dimensional array of read contents
   */
  public static SheetValues loadSheetValues(XSSFWorkbook workbook, String sheetName, int lastRowNum, int lastCellNum) {
    XSSFSheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      throw new RuntimeException("Sheet not found: " + sheetName);
    }

    int lastRowIndex = lastRowNum;
    if (lastRowIndex == 0) {
      lastRowIndex = sheet.getLastRowNum();
    }
    SheetValues rows = new SheetValues();
    int emptyRows = 0;
    Cell cell;
    for (int i = 0; i <= lastRowIndex; i++) {
      SheetRow row = new SheetRow();
      XSSFRow xssRow = sheet.getRow(i);

      int lastCellIndex = lastCellNum;

      if (xssRow == null) {
        if (lastCellNum > 0) {
          for (int j = 0; j < lastCellIndex; j++) {
            cell = parseCell(null);
            row.add(cell);
          }
        }
        rows.add(row);
        emptyRows++;
        continue;
      }

      if (lastCellNum == 0) {
        lastCellIndex = xssRow.getLastCellNum();
      }

      int valExists = 0;
      for (int j = 0; j < lastCellIndex; j++) {
        XSSFCell xssFcell = xssRow.getCell(j);
        cell = parseCell(xssFcell);
        row.add(cell);
        if ((xssFcell != null) && (!"".equals(cell.getValue()))) {
          valExists++;
        }
      }
      rows.add(row);

      // If there is no value in Row, it will be counted as a useless row.
      if (valExists == 0) {
        emptyRows++;
      } else {
        emptyRows = 0;
      }
    }

    // The empty part of Last is useless, so delete it.
    for (int i = 0; i < emptyRows; i++) {
      int removeIndex = rows.size() - 1;
      rows.remove(removeIndex);
    }

    return rows;
  }

  private static Cell parseCell(XSSFCell xssFcell) {
    Cell cell = new Cell();

    if (xssFcell == null) {
      cell.setXssFcell(null);
      cell.setValue("");
      return cell;
    }

    cell.setXssFcell(xssFcell);

    HSSFDataFormatter hdf = new HSSFDataFormatter();
    String cellString = hdf.formatCellValue(xssFcell);
    CellType cellType = xssFcell.getCellType();
    if ((CellType.FORMULA).equals(cellType)) {
      cellString = xssFcell.getRawValue();
      String formula = xssFcell.getCellFormula();
      cell.setFormula(formula);
    }
    cell.setValue(cellString);

    XSSFCellStyle style = xssFcell.getCellStyle();

    XSSFColor bgColor = style.getFillForegroundXSSFColor();
    String bgColorRGBHex = getRGBHex(bgColor);
    cell.setBackgroundColorRGBHex(bgColorRGBHex);

    XSSFFont font = style.getFont();
    XSSFColor fontColor = font.getXSSFColor();
    String fontColorRGBHex = getRGBHex(fontColor);
    cell.setFontColorRGBHex(fontColorRGBHex);

    return cell;
  }

  private static String getRGBHex(XSSFColor color) {
    String rgbHex = null;
    if (color != null) {
      rgbHex = color.getARGBHex();
      rgbHex = rgbHex.substring(2);
    }
    return rgbHex;
  }

  /**
   * Returns the sheet names of the book.
   *
   * @param path
   *          the book file path
   * @return an array of the sheet names
   */
  public static String[] getSheetNames(String path) {
    File file = new File(path);
    return getSheetNames(file);
  }

  /**
   * Returns the sheet names of the book.
   *
   * @param file
   *          the book file
   * @return an array of the sheet names
   */
  public static String[] getSheetNames(File file) {
    Path path = file.toPath();
    String[] sheetNames;
    try (FileSystem fs = FileSystems.newFileSystem(path, null)) {
      XmlContentHandler xmlContentHandler = new XmlContentHandler();
      Path xmlPath = fs.getPath("xl/workbook.xml");

      try (InputStream is = Files.newInputStream(xmlPath)) {
        InputSource source = new InputSource(is);
        XMLReader parser = XMLReaderFactory.createXMLReader();
        parser.setContentHandler(xmlContentHandler);
        parser.parse(source);
      }

      List<String> names = xmlContentHandler.names;
      int nameCount = names.size();
      sheetNames = new String[nameCount];
      names.toArray(sheetNames);
    } catch (Exception e) {
      sheetNames = null;
    }
    return sheetNames;
  }

  private static class XmlContentHandler extends DefaultHandler {
    private final List<String> names;

    public XmlContentHandler() {
      names = new ArrayList<>();
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
      if ("sheet".equals(qName)) {
        String name = attributes.getValue("name");
        names.add(name);
      }
    }
  }

}
