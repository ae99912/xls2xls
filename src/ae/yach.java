/*
 * Copyright (c) 2023. AE
 */

/*
 * класс Ячейка,
 * храним номер строки и столбца, сформированных на основе названия ячейки
 *
 */
package ae;

import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

class yach {
  final static Pattern cell_pattern = Pattern.compile("([A-Z]+)([0-9]+)");  // паттерн для имени ячейки A12, B3 ...
  int irow;       // номер строки  1 - 1
  int icol;       // номер столбца A - 1
  String  name;   // имя ячейки (для справки, по программе не нужно)

  yach() {}

  /**
   * установить значения
   * @param iCol  колонка
   * @param iRow  строка
   * @param Name  имя
   */
  yach(int iRow,int iCol,  String Name)
  {
    this.irow = iRow;
    this.icol = iCol;
    this.name = Name;
  }


  /**
   * преобразовать строку с ячейкой(ами) в набор ячеек
   * @param kartCellStr  строка с ячейкой (A21) или диапазоном (C2:D40)
   * @return набор ячеек
   */
  public Set<yach> set(String kartCellStr)
  {
    String sss = kartCellStr.toUpperCase().replaceAll ("\\s", "");
    if( sss.length() < 1 )
      return null;
    //
    HashSet<yach> yset = new HashSet<>();
    try {
      Matcher mat = cell_pattern.matcher(sss);
      if (!mat.find()) {
        throw new NumberFormatException("not found cell name");
      }
      int c1 = getExcelColumnNumber(mat.group(1));
      int r1 = Integer.parseInt(mat.group(2));
      if (c1 < 1 || r1 < 1) {
        throw new NumberFormatException("number less 1");
      }
      this.icol = c1;
      this.irow = r1;
      this.name = kartCellStr;
      yset.add(this);
      if (!mat.find())
        return yset;
      // есть вторая ячейка, значит диапазон
      int c2 = getExcelColumnNumber(mat.group(1));
      int r2 = Integer.parseInt(mat.group(2));
      if (c2 < 1 || r2 < 1) {
        throw new NumberFormatException("number less 1");
      }
      if (c1 > c2) {
        int a = c1;  c1 = c2;  c2 = a;
      }
      if (r1 > r2) {
        int a = r1;  r1 = r2;  r2 = a;
      }
      for (int ic = c1; ic <= c2; ic++) {
        for (int jr = r1; jr <= r2; jr++) {
          yach yy = new yach(jr, ic, kartCellStr);
          yset.add(yy);
        }
      }
    } catch (Exception e) {
      System.err.println("?-Error-cell.setAll('" + kartCellStr + "') error conversion: " + e.getMessage());
      return null;
    }
    return yset;
  }

  @Override
  public boolean equals(Object o) {
    // https://javarush.com/groups/posts/2179-metodih-equals--hashcode-praktika-ispoljhzovanija
    if (getClass() != o.getClass())
      return false;
    yach yo = (yach)o;
    return this.icol == yo.icol  && this.irow == yo.irow;
  }

  @Override
  public  int hashCode() {
    // https://javarush.com/groups/posts/2179-metodih-equals--hashcode-praktika-ispoljhzovanija
    String s = this.icol + "," + this.irow;
    return s.hashCode();
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
