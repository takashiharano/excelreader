package com.libutil.excelreader.test.book;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.libutil.excelreader.ExcelLoader;
import com.libutil.excelreader.LoadingListener;
import com.libutil.excelreader.test.book.model.Sheet1ValuesMap;
import com.libutil.excelreader.test.book.parser.Sheet1Parser;

public class Book1 {

  private String bookPath1;

  // Loading listener of book File
  private LoadingListener loadingListener;

  // The model objects of the book/sheet
  private Sheet1ValuesMap sheet1ValuesMap;

  private String[] sheetNames;

  public Book1(String bookPath) throws IOException {
    this(bookPath, null);
  }

  public Book1(String bookPath1, LoadingListener loadingListener) throws IOException {
    this.bookPath1 = bookPath1;
    this.loadingListener = loadingListener;
    loadBookFiles();
  }

  private void loadBookFiles() throws IOException {
    // Book file existence check
    ArrayList<String> errFiles = new ArrayList<>();
    fileCheck(errFiles, bookPath1);
    if (errFiles.size() > 0) {
      String errMessage = buildFileErrorMessage(errFiles);
      throw new IOException(errMessage);
    }

    onLoadStart("Book1");
    loadBook1(bookPath1);
    onLoadComplete("Book1");

    // Processes what refers to information across the board here.
    postProcess();
  }

  private void fileCheck(ArrayList<String> errFiles, String path) {
    File file = new File(path);
    if (!file.exists()) {
      errFiles.add(path);
    }
  }

  private String buildFileErrorMessage(ArrayList<String> errFiles) {
    StringBuilder sb = new StringBuilder("File not found: ");
    for (int i = 0; i < errFiles.size(); i++) {
      if (i > 0) {
        sb.append(", ");
      }
      sb.append(errFiles.get(i));
    }
    return sb.toString();
  }

  private void loadBook1(String specFilePath) throws IOException {
    XSSFWorkbook workbook = ExcelLoader.openBook(specFilePath);

    Sheet1Parser parser = new Sheet1Parser();
    this.sheet1ValuesMap = parser.parse(workbook);

    this.sheetNames = ExcelLoader.getSheetNames(specFilePath);

    workbook.close();
  }

  public Sheet1ValuesMap getAllValues() {
    return sheet1ValuesMap;
  }

  public String[] getSheetNames() {
    return sheetNames;
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

  /**
   * Post process.
   */
  private void postProcess() {

  }

}
