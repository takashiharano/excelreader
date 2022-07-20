package com.libutil.excelreader.test.book;

import com.libutil.excelreader.LoadingListener;

public class ExcelLoadingListener implements LoadingListener {

  @Override
  public void onLoadStart(String name) {
    System.out.println(name + " load started");
  }

  @Override
  public void onLoadComplete(String name) {
    System.out.println(name + " load completed");
  }

}
