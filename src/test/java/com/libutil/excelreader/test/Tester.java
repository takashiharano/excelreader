package com.libutil.excelreader.test;

import java.io.IOException;
import java.util.Map.Entry;

import com.libutil.excelreader.LoadingListener;
import com.libutil.excelreader.test.book.Book1;
import com.libutil.excelreader.test.book.ExcelLoadingListener;
import com.libutil.excelreader.test.book.model.Sheet1Values;
import com.libutil.excelreader.test.book.model.Sheet1ValuesMap;

public class Tester {

  public static void main(String args[]) {
    String path = "C:/test/Book1.xlsx";
    LoadingListener listener = new ExcelLoadingListener();
    try {
      Book1 book1 = new Book1(path, listener);
      test(book1);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private static void test(Book1 book1) {
    dumpValues(book1);
  }

  private static void dumpValues(Book1 book1) {
    StringBuilder sb = new StringBuilder();
    Sheet1ValuesMap map = book1.getAllValues();

    for (Entry<String, Sheet1Values> entry : map.entrySet()) {
      Sheet1Values values = entry.getValue();
      int no = values.getNo();
      String key = values.getKey();
      String itemA = values.getItemA();
      String itemB = values.getItemB();
      String itemC = values.getItemC();
      String itemD = values.getItemD();

      sb.append("No=" + no);
      sb.append("\t");
      sb.append("key=" + key);
      sb.append("\t");
      sb.append("itemA=" + itemA);
      sb.append("\t");
      sb.append("itemB=" + itemB);
      sb.append("\t");
      sb.append("itemC=" + itemC);
      sb.append("\t");
      sb.append("itemD=" + itemD);
      sb.append("\n");
    }

    System.out.println(sb.toString());
  }

}
