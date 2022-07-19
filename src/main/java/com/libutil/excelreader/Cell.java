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
   * Returns the value of the cell as a double.
   *
   * @return The value of the cell
   */
  public double getDoubleValue() {
    double v = ExcelStringUtil.toDouble(value);
    return v;
  }

  /**
   * Returns the value of the cell as a double.
   *
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public double getDoubleValue(double defaultValue) {
    double v = ExcelStringUtil.toDouble(value, defaultValue);
    return v;
  }

  /**
   * Returns the value of the cell as a float.
   *
   * @return The value of the cell
   */
  public float getFloatValue() {
    float v = ExcelStringUtil.toFloat(value);
    return v;
  }

  /**
   * Returns the value of the cell as a float.
   *
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public float getFloatValue(float defaultValue) {
    float v = ExcelStringUtil.toFloat(value, defaultValue);
    return v;
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
   * Returns the value of the cell as an integer.
   *
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public int getIntValue(int defaultValue) {
    int v = ExcelStringUtil.toInteger(value, defaultValue);
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
   * Returns the value of the cell as a long.
   *
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public long getLongValue(long defaultValue) {
    long v = ExcelStringUtil.toLong(value, defaultValue);
    return v;
  }

  /**
   * Returns whether the cell value can be considered true.
   *
   * @return A zero value, "FALSE", "", null, are converted to false; any other
   *         value is converted to true. The value is case-insensitive.
   */
  public boolean isTrue() {
    if (value == null) {
      return false;
    }
    value = value.toUpperCase();
    if (value.equals("") || value.equals("0") || value.equals("FALSE")) {
      return false;
    }
    return true;
  }

  /**
   * Returns whether the cell value can be considered true.
   *
   * @param trueValue
   *          the string that is considered true
   * @return true if the value equals trueValue; false otherwise
   */
  public boolean isTrue(String trueValue) {
    if (value == null) {
      return false;
    }
    return value.equals(trueValue);
  }

  /**
   * Returns whether the cell value can be considered true.
   *
   * @param trueValues
   *          the values to be true
   * @return true if the cell value is equal to one of trueValues.
   */
  public boolean isTrue(String[] trueValues) {
    for (int i = 0; i < trueValues.length; i++) {
      String v = trueValues[i];
      if (v == null) {
        if (value == null) {
          return true;
        } else {
          continue;
        }
      }
      if (v.equals(value)) {
        return true;
      }
    }
    return false;
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
   * Returns if the cell has font color.
   *
   * @return true if the cell has font color
   */
  public boolean hasFontColor() {
    if ("000000".equals(fontColorRGBHex)) {
      return false;
    }
    return true;
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

  /**
   * Returns if the cell has background color.
   *
   * @return true if the cell has background color
   */
  public boolean hasBackgroundColor() {
    if (backgroundColorRGBHex == null) {
      return false;
    }
    return true;
  }

}
