package com.libutil.excelreader.test.book;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.libutil.excelreader.ExcelLoader;
import com.libutil.excelreader.LoadingListener;
import com.libutil.excelreader.test.book.model.Sheet1RecordMap;
import com.libutil.excelreader.test.book.parser.Sheet1Parser;

public class Main {

  private String bookPath;;

  // Loading listener of specification File
  private LoadingListener loadingListener;

  // The model objects of the specifications
  private Sheet1RecordMap sheet1RecordMap;

  public Main(String bookPath) throws IOException {
    this(bookPath, null);
  }

  public Main(String bookPath, LoadingListener loadingListener) throws IOException {
    this.bookPath = bookPath;
    this.loadingListener = loadingListener;
    loadSpecifications();
  }

  private void loadSpecifications() throws IOException {
    // Specification file existence check
    ArrayList<String> errFiles = new ArrayList<>();
    fileCheck(errFiles, bookPath);
    if (errFiles.size() > 0) {
      String errMessage = buildFileErrorMessage(errFiles);
      throw new IOException(errMessage);
    }

    onLoadStart("Book1");
    loadBook1(bookPath);
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

    Sheet1Parser sheet1Parser = new Sheet1Parser();
    this.sheet1RecordMap = sheet1Parser.parse(workbook);

    workbook.close();
  }

  public Sheet1RecordMap getAllRecords() {
    return sheet1RecordMap;
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
