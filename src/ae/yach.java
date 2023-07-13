/*
 * Copyright (c) 2023. AE
 */

/*
 * класс Ячейка,
 * храним номер строки и столбца, сформированных на основе названия ячейки
 *
 */
package ae;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

class yach {
  // паттерн для имени ячейки A12, B3 ...
  final static Pattern cell_pattern = Pattern.compile("([A-Z]+)([0-9]+)");
  int irow;       // номер строки  1 - 1
  int icol;       // номер столбца A - 1
  String  name;   // имя ячейки (для справки, по программе не нужно)

  /**
   * установить номера строки и столбца по строке с именем ячейки
   * @param kartCellStr - название ячейки (A1, B12 и т.д.)
   * @return  значение установлено
   */
  boolean set(String kartCellStr)
  {
    try {
      Matcher mat = cell_pattern.matcher(kartCellStr.toUpperCase());
      if(!mat.find()) {
        throw new NumberFormatException("not found cell name");
      }
      int c = getExcelColumnNumber(mat.group(1));
      int r = Integer.parseInt(mat.group(2));
      if( c < 1 || r < 1 ) {
        throw new NumberFormatException("number less 1");
      }
      this.icol = c;
      this.irow = r;
      this.name = kartCellStr;  // справочная инфа
    } catch (Exception e) {
      System.err.println("?-Error-cell.set('" + kartCellStr + "') error conversion: " + e.getMessage());
      return false;
    }
    return true;
  }

  /**
   * преобразовать имя столбца Excel в его номер
   * @param column - строка имени столбца
   * @return номер столбца
   */
  private static int getExcelColumnNumber(String column)
  {
    int result = 0;
    for(int i = 0; i < column.length(); i++) {
      result *= 26;
      result += column.charAt(i) - 'A' + 1;
    }
    return result;
  }

} // end of class
