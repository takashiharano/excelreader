package com.libutil.excelreader.test.book.parser;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.libutil.excelreader.ExcelLoader;
import com.libutil.excelreader.SheetRow;
import com.libutil.excelreader.SheetValues;
import com.libutil.excelreader.test.book.model.Sheet1Values;
import com.libutil.excelreader.test.book.model.Sheet1ValuesMap;

/**
 * The Sheet1 parser.
 */
public class Sheet1Parser {

  // Sheet structure
  public static final String SHEET_NAME = "Sheet1";
  public static final int DEFINITION_START_ROW = 2;
  public static final int DEFINITION_LAST_ROW = 0; // Auto detection
  public static final String COL_INDEX_NO = "A";
  public static final String COL_INDEX_KEY = "B";
  public static final String COL_INDEX_ITEM_A = "C";
  public static final String COL_INDEX_ITEM_B = "D";
  public static final String COL_INDEX_ITEM_C = "E";
  public static final String LAST_COL_INDEX = "E";

  /**
   * Parses the Sheet1
   *
   * @param workbook
   *          the Excel book object
   * @return the values object map
   */
  public Sheet1ValuesMap parse(XSSFWorkbook workbook) {
    // Open the book and read the values.
    SheetValues sheetRows = ExcelLoader.loadSheetValues(workbook, SHEET_NAME, DEFINITION_LAST_ROW, LAST_COL_INDEX);

    Sheet1ValuesMap valuesMap = new Sheet1ValuesMap();

    // Loop the row of Sheet
    for (int i = DEFINITION_START_ROW; i <= sheetRows.size(); i++) {
      SheetRow row = sheetRows.getRow(i);
      Sheet1Values values = parseRow(row);

      String key = values.getKey();
      valuesMap.put(key, values);
    }

    return valuesMap;
  }

  /**
   * Parses the Row and stores the information in the object.
   *
   * @param row
   *          Array of values in Row
   * @return values object
   */
  private Sheet1Values parseRow(SheetRow row) {
    String strNo = row.getValue(COL_INDEX_NO);
    if ((strNo == null) || strNo.trim().equals("")) {
      return null;
    }

    Sheet1Values values = new Sheet1Values();

    // COL: No
    int no = row.getIntValue(COL_INDEX_NO);
    values.setNo(no);

    // COL: ItemA
    String key = row.getValue(COL_INDEX_KEY);
    values.setKey(key);

    // COL: ItemA
    String itemA = row.getValue(COL_INDEX_ITEM_A);
    values.setItemA(itemA);

    // COL: ItemB
    String itemB = row.getValue(COL_INDEX_ITEM_B);
    values.setItemB(itemB);

    // COL: ItemC
    String itemC = row.getValue(COL_INDEX_ITEM_C);
    values.setItemC(itemC);

    return values;
  }

}
