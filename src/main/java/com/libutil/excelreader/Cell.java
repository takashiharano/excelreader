package com.libutil.excelreader;

import org.apache.poi.xssf.usermodel.XSSFCell;

public class Cell {

  public static final int CELL_TYPE_NONE = 0;
  public static final int CELL_TYPE_BLANK = 1;
  public static final int CELL_TYPE_BOOLEAN = 2;
  public static final int CELL_TYPE_ERROR = 3;
  public static final int CELL_TYPE_FORMULA = 4;
  public static final int CELL_TYPE_NUMERIC = 5;
  public static final int CELL_TYPE_STRING = 6;

  private XSSFCell xssFcell;
  private int cellType;
  private String value;
  private String formula;
  private String fontColorRGBHex;
  private String backgroundColorRGBHex;

  /**
   * Returns the XSSFCell object.
   *
   * @return the XSSFCell object
   */
  public XSSFCell getXssFcell() {
    return xssFcell;
  }

  /**
   * Sets the XSSFCell object.
   *
   * @param xssFcell
   *          the XSSFCell object
   */
  public void setXssFcell(XSSFCell xssFcell) {
    this.xssFcell = xssFcell;
  }

  /**
   * Returns cell type by number.
   *
   * @return the cell type. 0-6
   */
  public int getCellType() {
    return cellType;
  }

  public void setCellType(int cellType) {
    this.cellType = cellType;
  }

  /**
   * Returns the cell value.
   *
   * @return the value
   */
  public String getValue() {
    return value;
  }

  /**
   * Sets the cell value.
   *
   * @param value
   *          the value
   */
  public void setValue(String value) {
    this.value = value;
  }

  /**
   * Returns the value of the cell as an integer.
   *
   * @return The value of the cell
   */
  public int getIntValue() {
    int v = ExcelStringUtil.toInteger(value);
    return v;
  }

  /**
   * Returns the value of the cell as a long.
   *
   * @return The value of the cell
   */
  public long getLongValue() {
    long v = ExcelStringUtil.toLong(value);
    return v;
  }

  /**
   * Returns if the cell value at the column position is empty.
   *
   * @return true if the value is empty
   */
  public boolean isEmpty() {
    if ((value == null) || ("".equals(value))) {
      return true;
    }
    return false;
  }

  /**
   * Returns a formula for the cell, for example, SUM(C4:E4).
   *
   * @return a formula for the cell
   */
  public String getFormula() {
    return formula;
  }

  public void setFormula(String formula) {
    this.formula = formula;
  }

  /**
   * Returns if the cell has a formula.
   *
   * @return true if the cell has a formula
   */
  public boolean hasFormula() {
    return !(formula == null);
  }

  /**
   * Returns the font RGB value in hex format.
   *
   * @return the RGB value in hex string format, eg FF0000.
   */
  public String getFontColorRGBHex() {
    return fontColorRGBHex;
  }

  /**
   * Sets the font RGB value from hex format.
   *
   * @param fontColorRGBHex
   *          color RGB hex string
   */
  public void setFontColorRGBHex(String fontColorRGBHex) {
    this.fontColorRGBHex = fontColorRGBHex;
  }

  /**
   * Returns the background RGB value in hex format.
   *
   * @return the RGB value in hex string format, eg FF0000.
   */
  public String getBackgroundColorRGBHex() {
    return backgroundColorRGBHex;
  }

  /**
   * Sets the background RGB value from hex format.
   * 
   * @param backgroundColorRGBHex
   *          color RGB hex string
   */
  public void setBackgroundColorRGBHex(String backgroundColorRGBHex) {
    this.backgroundColorRGBHex = backgroundColorRGBHex;
  }

}
