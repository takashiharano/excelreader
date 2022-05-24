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
 * This class treats the value of the entire Excel Sheet as if it were stored.
 */
public class SheetValues extends ArrayList<SheetRow> {

  private static final long serialVersionUID = 1L;

  /**
   * Returns the data for the row specified by rowIndex.
   *
   * @param rowIndex
   *          The index of row (1-1048576)
   * @return row data
   */
  public SheetRow getRow(int rowIndex) {
    return get(rowIndex - 1);
  }

  /**
   * Returns the cell object corresponding to the position specified by col and
   * row.
   *
   * @param col
   *          The position of column (A-XFD)
   * @param row
   *          The index of row (1-1048576)
   * @return The cell object
   */
  public Cell getCell(String col, int row) {
    int colIndex = ExcelStringUtil.xlscol(col);
    return getCell(colIndex, row);
  }

  /**
   * Returns the cell object corresponding to the position specified by col and
   * row.
   *
   * @param col
   *          The position of column (1-16384)
   * @param row
   *          The index of row (1-1048576)
   * @return The cell object
   */
  public Cell getCell(int col, int row) {
    SheetRow cols = get(row - 1);
    return cols.getCell(col);
  }

  /**
   * Returns the value of the cell corresponding to the position specified by col
   * and row.
   *
   * @param col
   *          The position of column (A-XFD)
   * @param row
   *          The index of row (1-1048576)
   * @return The value of the cell
   */
  public String getValue(String col, int row) {
    Cell cell = getCell(col, row);
    return cell.getValue();
  }

  /**
   * Returns the value of the cell corresponding to the position specified by col
   * and row.
   *
   * @param col
   *          The position of column (1-16384)
   * @param row
   *          The index of row (1-1048576)
   * @return The value of the cell
   */
  public String getValue(int col, int row) {
    Cell cell = getCell(col, row);
    return cell.getValue();
  }

  /**
   * Returns the value of the cell corresponding to the position specified by col
   * and row as an integer.
   *
   * @param col
   *          The position of column (A-XFD)
   * @param row
   *          The index of row (1-1048576)
   * @return The value of the cell
   */
  public int getIntValue(String col, int row) {
    Cell cell = getCell(col, row);
    return cell.getIntValue();
  }

  /**
   * Returns the value of the cell corresponding to the position specified by col
   * and row as an integer.
   *
   * @param col
   *          The position of column (1-16384)
   * @param row
   *          The index of row (1-1048576)
   * @return The value of the cell
   */
  public int getIntValue(int col, int row) {
    Cell cell = getCell(col, row);
    return cell.getIntValue();
  }

  /**
   * Returns the value of the cell corresponding to the position specified by col
   * and row as a long.
   *
   * @param col
   *          The position of column (A-XFD)
   * @param row
   *          The index of row (1-1048576)
   * @return The value of the cell
   */
  public long getLongValue(String col, int row) {
    Cell cell = getCell(col, row);
    return cell.getLongValue();
  }

  /**
   * Returns the value of the cell corresponding to the position specified by col
   * and row as a long.
   *
   * @param col
   *          The position of column (1-16384)
   * @param row
   *          The index of row (1-1048576)
   * @return The value of the cell
   */
  public long getLongValue(int col, int row) {
    Cell cell = getCell(col, row);
    return cell.getLongValue();
  }

  /**
   * Returns if the cell value at the column position is empty.
   *
   * @param col
   *          The position of column (A-XFD)
   * @param row
   *          The index of row (1-1048576)
   * @return true if the value is empty
   */
  public boolean isEmpty(String col, int row) {
    Cell cell = getCell(col, row);
    return cell.isEmpty();
  }

  /**
   * Returns if the cell value at the column position is empty.
   *
   * @param col
   *          The position of column (1-16384)
   * @param row
   *          The index of row (1-1048576)
   * @return true if the value is empty
   */
  public boolean isEmpty(int col, int row) {
    Cell cell = getCell(col, row);
    return cell.isEmpty();
  }

}
