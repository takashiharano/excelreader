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

import java.util.ArrayList;

/**
 * This class stores the data for one line of Excel.
 */
public class SheetRow extends ArrayList<Cell> {

  private static final long serialVersionUID = 1L;

  /**
   * Returns the cell object at the column position.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return The cell object
   */
  public Cell getCell(String colIndex) {
    int index = ExcelStringUtil.xlscol(colIndex);
    return getCell(index);
  }

  /**
   * Returns the cell object at the column position.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return The cell object
   */
  public Cell getCell(int colIndex) {
    int index = colIndex - 1;
    Cell cell;
    try {
      cell = get(index);
    } catch (Exception e) {
      int listSize = this.size();
      String strCol = ExcelStringUtil.xlscol(colIndex);
      String lastCol = ExcelStringUtil.xlscol(listSize);
      String msg = "Get cell error: col=" + colIndex + "(" + strCol + ") : last col=" + listSize + "(" + lastCol + ")";
      throw new RuntimeException(msg, e);
    }
    return cell;
  }

  /**
   * Returns the cell value at the column position.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return the value of the cell
   */
  public String getValue(String colIndex) {
    int index = ExcelStringUtil.xlscol(colIndex);
    return getValue(index);
  }

  /**
   * Returns the cell value at the column position.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return The value of the cell
   */
  public String getValue(int colIndex) {
    Cell cell = getCell(colIndex);
    String value = cell.getValue();
    return value;
  }

  /**
   * Returns the cell value at the column position as a double.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return The value of the cell
   */
  public double getDoubleValue(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getDoubleValue();
  }

  /**
   * Returns the cell value at the column position as a double.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return The value of the cell
   */
  public double getDoubleValue(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getDoubleValue();
  }

  /**
   * Returns the cell value at the column position as a double.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public double getDoubleValue(String colIndex, double defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getDoubleValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as a double.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public double getDoubleValue(int colIndex, double defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getDoubleValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as a float.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return The value of the cell
   */
  public float getFloatValue(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getFloatValue();
  }

  /**
   * Returns the cell value at the column position as a float.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return The value of the cell
   */
  public float getFloatValue(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getFloatValue();
  }

  /**
   * Returns the cell value at the column position as a float.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public float getFloatValue(String colIndex, float defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getFloatValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as a float.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public float getFloatValue(int colIndex, float defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getFloatValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as an integer.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return The value of the cell
   */
  public int getIntValue(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getIntValue();
  }

  /**
   * Returns the cell value at the column position as an integer.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return The value of the cell
   */
  public int getIntValue(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getIntValue();
  }

  /**
   * Returns the cell value at the column position as an integer.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public int getIntValue(String colIndex, int defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getIntValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as an integer.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public int getIntValue(int colIndex, int defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getIntValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as a long.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return The value of the cell
   */
  public long getLongValue(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getLongValue();
  }

  /**
   * Returns the cell value at the column position as a long.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return The value of the cell
   */
  public long getLongValue(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getLongValue();
  }

  /**
   * Returns the cell value at the column position as a long.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public long getLongValue(String colIndex, long defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getLongValue(defaultValue);
  }

  /**
   * Returns the cell value at the column position as a long.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @param defaultValue
   *          value for parse error
   * @return The value of the cell
   */
  public long getLongValue(int colIndex, long defaultValue) {
    Cell cell = getCell(colIndex);
    return cell.getLongValue(defaultValue);
  }

  /**
   * Returns whether the cell value at the column position can be considered true.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return A zero value, "FALSE", "", null, are converted to false; any other
   *         value is converted to true. The value is case-insensitive.
   */
  public boolean isTrue(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.isTrue();
  }

  /**
   * Returns whether the cell value at the column position can be considered true.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return A zero value, "FALSE", "", null, are converted to false; any other
   *         value is converted to true. The value is case-insensitive.
   */
  public boolean isTrue(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.isTrue();
  }

  /**
   * Returns whether the cell value at the column position can be considered true.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @param trueValue
   *          the string that is considered true
   * @return true if the value is equal to trueValue; false otherwise
   */
  public boolean isTrue(String colIndex, String trueValue) {
    Cell cell = getCell(colIndex);
    return cell.isTrue(trueValue);
  }

  /**
   * Returns whether the cell value at the column position can be considered true.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @param trueValue
   *          the string that is considered true
   * @return true if the value is equal to trueValue; false otherwise
   */
  public boolean isTrue(int colIndex, String trueValue) {
    Cell cell = getCell(colIndex);
    return cell.isTrue(trueValue);
  }

  /**
   * Returns whether the cell value at the column position can be considered true.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @param trueValues
   *          the values to be true
   * @return true if the value is equal to one of trueValues.
   */
  public boolean isTrue(String colIndex, String[] trueValues) {
    Cell cell = getCell(colIndex);
    return cell.isTrue(trueValues);
  }

  /**
   * Returns whether the cell value at the column position can be considered true.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @param trueValues
   *          the values to be true
   * @return true if the value is equal to one of trueValues.
   */
  public boolean isTrue(int colIndex, String[] trueValues) {
    Cell cell = getCell(colIndex);
    return cell.isTrue(trueValues);
  }

  /**
   * Returns the font RGB value in hex format.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return the RGB value in hex string format, eg FF0000.
   */
  public String getFontColorRGBHex(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getFontColorRGBHex();
  }

  /**
   * Returns the font RGB value in hex format.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return the RGB value in hex string format, eg FF0000.
   */
  public String getFontColorRGBHex(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getFontColorRGBHex();
  }

  /**
   * Returns if the cell has font color.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return true if the cell has font color
   */
  public boolean hasFontColor(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.hasFontColor();
  }

  /**
   * Returns if the cell has font color.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return true if the cell has font color
   */
  public boolean hasFontColor(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.hasFontColor();
  }

  /**
   * Returns the background RGB value in hex format.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return the RGB value in hex string format, eg FF0000.
   */
  public String getBackgroundColorRGBHex(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getBackgroundColorRGBHex();
  }

  /**
   * Returns the background RGB value in hex format.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return the RGB value in hex string format, eg FF0000.
   */
  public String getBackgroundColorRGBHex(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.getBackgroundColorRGBHex();
  }

  /**
   * Returns if the cell has background color.
   *
   * @param colIndex
   *          The column position (1-16384)
   * 
   * @return true if the cell has background color
   */
  public boolean hasBackgroundColor(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.hasBackgroundColor();
  }

  /**
   * Returns if the cell has background color.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * 
   * @return true if the cell has background color
   */
  public boolean hasBackgroundColor(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.hasBackgroundColor();
  }

  /**
   * Returns if the cell value at the column position is empty.
   *
   * @param colIndex
   *          The column position (A-XFD)
   * @return true if the value is empty
   */
  public boolean isEmpty(String colIndex) {
    Cell cell = getCell(colIndex);
    return cell.isEmpty();
  }

  /**
   * Returns if the cell value at the column position is empty.
   *
   * @param colIndex
   *          The column position (1-16384)
   * @return true if the value is empty
   */
  public boolean isEmpty(int colIndex) {
    Cell cell = getCell(colIndex);
    return cell.isEmpty();
  }

}
