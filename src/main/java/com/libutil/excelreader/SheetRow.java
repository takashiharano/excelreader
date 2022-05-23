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
      String lastCol = ExcelStringUtil.xlscol(listSize);
      throw new RuntimeException("Get cell error: col=" + colIndex + " last col=" + lastCol, e);
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
    int index = colIndex - 1;
    Cell cell = getCell(index);
    String value = cell.getValue();
    return value;
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
