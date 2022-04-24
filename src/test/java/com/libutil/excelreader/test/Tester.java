package com.libutil.excelreader.test;

import java.io.IOException;
import java.util.Map.Entry;

import com.libutil.excelreader.LoadingListener;
import com.libutil.excelreader.test.book.ExcelLoadingListener;
import com.libutil.excelreader.test.book.Main;
import com.libutil.excelreader.test.book.model.Sheet1Record;
import com.libutil.excelreader.test.book.model.Sheet1RecordMap;

public class Tester {

  public static void main(String args[]) {
    LoadingListener listener = new ExcelLoadingListener();
    String bookPath = "C:/test/Book1.xlsx";
    try {
      Main m = new Main(bookPath, listener);
      test(m);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private static void test(Main m) {
    StringBuilder sb = new StringBuilder();
    Sheet1RecordMap map = m.getAllRecords();

    for (Entry<String, Sheet1Record> entry : map.entrySet()) {
      Sheet1Record record = entry.getValue();
      int no = record.getNo();
      String key = record.getKey();
      String itemA = record.getItemA();
      String itemB = record.getItemB();
      String itemC = record.getItemC();

      sb.append("No=" + no);
      sb.append("\t");
      sb.append("key=" + key);
      sb.append("\t");
      sb.append("itemA=" + itemA);
      sb.append("\t");
      sb.append("itemB=" + itemB);
      sb.append("\t");
      sb.append("itemC=" + itemC);
      sb.append("\n");
    }

    System.out.println(sb.toString());
  }

}
