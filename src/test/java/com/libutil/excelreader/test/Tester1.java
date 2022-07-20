package com.libutil.excelreader.test;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.libutil.excelreader.ExcelLoader;
import com.libutil.excelreader.LoadingListener;
import com.libutil.excelreader.SheetRow;
import com.libutil.excelreader.SheetValues;
import com.libutil.excelreader.test.book.ExcelLoadingListener;

public class Tester1 {

  // Sheet structure
  public static final String SHEET_NAME = "Sheet1";
  public static final int DEFINITION_START_ROW = 1;
  public static final int DEFINITION_LAST_ROW = 0; // Auto detection
  public static final String LAST_COL_INDEX = null; // Auto detection

  LoadingListener loadingListener;

  public static void main(String args[]) {
    Tester1 tester = new Tester1();
    tester.test();
  }

  public void test() {
    String path = "C:/test/Book1.xlsx";
    loadingListener = new ExcelLoadingListener();
    try {
      _test(path);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private void _test(String path) throws IOException {
    onLoadStart("book");
    XSSFWorkbook workbook = ExcelLoader.openBook(path);
    parse(workbook);
    workbook.close();
    onLoadComplete("book");
  }

  private void onLoadStart(String name) {
    if (loadingListener != null) {
      loadingListener.onLoadStart(name);
    }
  }

  private void onLoadComplete(String name) {
    if (loadingListener != null) {
      loadingListener.onLoadComplete(name);
    }
  }

  public void parse(XSSFWorkbook workbook) {
    // Open the book and read the values.
    SheetValues sheetRows = ExcelLoader.loadSheetValues(workbook, SHEET_NAME, DEFINITION_LAST_ROW, LAST_COL_INDEX);

    for (int i = DEFINITION_START_ROW; i <= sheetRows.size(); i++) {
      SheetRow row = sheetRows.getRow(i);
      String values = parseRow(row);
      System.out.println(values);
    }
  }

  private String parseRow(SheetRow row) {
    StringBuilder sb = new StringBuilder();
    for (int i = 1; i <= row.size(); i++) {
      if (i > 1) {
        sb.append("\t");
      }
      String v = row.getValue(i);
      sb.append(v);
    }
    return sb.toString();
  }

}
