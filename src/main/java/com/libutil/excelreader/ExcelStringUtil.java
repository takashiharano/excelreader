/*
 * The MIT License
 *
 * Copyright 2022 Takashi Harano
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
 * 
 * ! This class is a subset of StrUtil on libutil.com.
 */
package com.libutil.excelreader;

import java.util.ArrayList;

/**
 * This class implements the string related processing.
 */
public class ExcelStringUtil {

  /**
   * Converts a string to a signed decimal integer.
   *
   * @param s
   *          a string value
   * @return an integer value
   */
  public static int toInteger(String s) {
    return toInteger(s, 0);
  }

  /**
   * Converts a string to a signed decimal integer.
   *
   * @param s
   *          a string value
   * @param defaultValue
   *          value for parse error
   * @return an integer value
   */
  public static int toInteger(String s, int defaultValue) {
    if (s == null) {
      return defaultValue;
    }
    s = s.trim();
    int v = defaultValue;
    try {
      v = Integer.parseInt(s);
    } catch (NumberFormatException e) {
      // nop
    }
    return v;
  }

  /**
   * Converts the Excel column letter to a numeric index.<br>
   * <br>
   * e.g., "A"=1, "B"=2, ... "Z"=26, "AA"=27, ... "XFD"=16384
   *
   * @param s
   *          the letter to convert ("A"-"XFD")
   * @return Index corresponding to a character (1-16384)
   */
  public static int xlscol(String s) {
    return xlscol(s, 0);
  }

  /**
   * Converts the Excel column letter to a numeric index.<br>
   * <br>
   * e.g., "A"=1, "B"=2, ... "Z"=26, "AA"=27, ... "XFD"=16384
   *
   * @param s
   *          the letter to convert ("A"-"XFD")
   * @param offset
   *          offset
   * @return Index corresponding to a character (1-16384)
   */
  public static int xlscol(String s, int offset) {
    return (int) Permutation.getIndex("ABCDEFGHIJKLMNOPQRSTUVWXYZ", s.toUpperCase()) + offset;
  }

  /**
   * Converts the Excel column index to a letter.<br>
   * <br>
   * e.g., 1="A", 2="B", ... 26="Z", 27="AA", ... 16384="XFD"
   *
   * @param n
   *          the index to convert (1-16384)
   * @return the letter corresponding to the index ("A-"XFD")
   */
  public static String xlscol(int n) {
    return Permutation.getString("ABCDEFGHIJKLMNOPQRSTUVWXYZ", n);
  }

  /**
   * String permutation.
   */
  public static class Permutation {
    /**
     * Count the number of total permutation patterns of the table.
     *
     * @param chars
     *          characters to use
     * @param length
     *          the length
     * @return the number of total pattern
     */
    public static long countTotal(String chars, int length) {
      String[] tbl = chars.split("");
      int c = tbl.length;
      int n = 0;
      for (int i = 1; i <= length; i++) {
        n += Math.pow(c, i);
      }
      return n;
    }

    /**
     * Returns the characters permutation index of the given pattern.
     *
     * @param chars
     *          characters to use
     * @param pattern
     *          a string
     * @return the index
     */
    public static long getIndex(String chars, String pattern) {
      int len = pattern.length();
      int rdx = chars.length();
      long idx = 0;
      for (int i = 0; i < len; i++) {
        int d = len - i - 1;
        String c = pattern.substring(d, d + 1);
        int v = chars.indexOf(c) + 1;
        long n = v * (long) Math.pow(rdx, i);
        idx += n;
      }
      return idx;
    }

    /**
     * Returns the string that appear in the specified order within the permutation
     * of the characters.
     *
     * @param chars
     *          characters to use
     * @param index
     *          the index
     * @return the string
     */
    public static String getString(String chars, long index) {
      if (index <= 0) {
        return "";
      }
      String[] tbl = chars.split("");
      int len = tbl.length;
      ArrayList<Integer> a = new ArrayList<>();
      a.add(-1);
      for (int i = 0; i < index; i++) {
        int j = 0;
        boolean cb = true;
        while (j < a.size()) {
          if (cb) {
            a.set(j, a.get(j) + 1);
            if (a.get(j) > len - 1) {
              a.set(j, 0);
              if (a.size() <= j + 1) {
                a.add(-1);
              }
            } else {
              cb = false;
            }
          }
          j++;
        }
      }
      int strLen = a.size();
      StringBuilder sb = new StringBuilder(strLen);
      for (int i = strLen - 1; i >= 0; i--) {
        sb.append(tbl[a.get(i)]);
      }
      return sb.toString();
    }
  }

}
