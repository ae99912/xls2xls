/*
 * Copyright (c) 2023. AE
 */
/*
  Карта переноса
  значений ячеек из входного файла Excel в выходной
  ячейки задаются в текстовом файле по одной ячейке в строке
  строка, начинающаяся с # - комментарий

 */
package ae;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class karta {
  // паттерн для имени ячейки
  final private Pattern cell_pattern = Pattern.compile("([A-Z]+)([0-9]+)");  // паттерн для имени ячейки A12, B3 ...

  HashSet<yach> f_set;    // набор множества ячеек

  /**
   * открыть и прочитать файл с картой и запомнить ячейки в множестве ячеек yach
   * @param fileName  имя файла
   * @return набор (множество) ячеек карты
   */
  Set<yach> open(String fileName)
  {
    try {
      f_set = new HashSet<>();
      BufferedReader rdr = new BufferedReader(new FileReader(fileName));
      String str;
      while( (str = rdr.readLine()) != null ) {
        if( str.length() > 1 && str.charAt(0) != '#' ) {
          addStr(str);
        }
      }
      //
    } catch (Exception e) {
      System.err.println("?-Error-karta.open('" + fileName + "') " + e.getMessage());
      return null;
    }
    return f_set;
  }

  /**
   * добавить в множество ячеек ячейки из строки карты как имя отдельной ячейки или диапазона ячеек
   * @param strKart   строка карты переноса
   */
  private void addStr(String strKart)
  {

    String sss = strKart.toUpperCase().replaceAll ("\\s", "");
    if( sss.length() < 1 )
      return;
    //
    try {
      Matcher mat = cell_pattern.matcher(sss);
      if (!mat.find()) {
        throw new NumberFormatException("not found cell name");
      }
      int c1 = getExcelColumnNumber(mat.group(1));
      int r1 = Integer.parseInt(mat.group(2));
      // добавим первую ячейку, неважно одна или диапазон
      this.f_set.add(new yach(r1, c1, strKart));
      //
      // проверим - есть еще ячейка в строке, если есть значит диапазон
      if (!mat.find())
        return;
      // есть вторая ячейка, значит диапазон
      int c2 = getExcelColumnNumber(mat.group(1));
      int r2 = Integer.parseInt(mat.group(2));
      // при задании диапазона правая граница д.б. больше левой
      if( c2 < c1 || r2 < r1 ) {
        throw new NumberFormatException("right less that left");
      }
      // заполним диапазон от края до края
      for (int ic = c1; ic <= c2; ic++) {
        for (int jr = r1; jr <= r2; jr++) {
          // добавим ячейку в набор
          this.f_set.add(new yach(jr, ic, strKart));
        }
      }
    } catch (Exception e) {
      System.err.println("?-Error-" + getClass() +".addStrKart('" + strKart + "') error conversion: " + e.getMessage());
    }
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
