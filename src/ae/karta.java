/*
 * Copyright (c) 2023. AE
 */
/*
  Карта переноса
  значений ячеек из входного файла Excel в выходной
  ячейки задаются в текстовом файле по одной ячейке в строке
  строка, начинающаяся с # - комментарий

  23.10.23 добавим к карте свойство (prop), пусть они начинаются с @
           первым свойством будет @only01 - читать только 0 и 1
  22.08.24 свойства назначаются следующим за ними ячейкам
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

  HashSet<yach> f_set;      // набор множества ячеек

  karta()
  {
    f_set  = new HashSet<>();
  }

  /**
   * открыть и прочитать файл с картой и запомнить ячейки в множестве ячеек yach
   * @param fileName  имя файла
   * @return набор (множество) ячеек карты
   */
  Set<yach> open(String fileName)
  {
    try {
     BufferedReader rdr = new BufferedReader(new FileReader(fileName));
      String str;
      String curProp = "";   // текущее свойство ячеек
      while( (str = rdr.readLine()) != null ) {
        if(str.length() > 1) {
          switch (str.charAt(0)) {
            case '#':   // комментарий
              break;

            case '@':   // свойство
              curProp = str.substring(1); // строка свойств
              break;

            default:    // ячейка
              addStr(str, curProp);
              break;
          }
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
  private void addStr(String strKart, String prop)
  {
    // буквы в верхний регистр и уберем все пробелы
    String sss = strKart.toUpperCase().replaceAll("\\s", "");
    if( sss.length() < 1 )
      return;
    //
    try {
      Matcher mat = cell_pattern.matcher(sss);
      if( !mat.find() ) {
        throw new NumberFormatException("not found cell name");
      }
      int c1 = getExcelColumnNumber(mat.group(1));
      int r1 = Integer.parseInt(mat.group(2));
      // добавим первую ячейку, неважно одна или диапазон
      this.f_set.add(new yach(r1, c1, prop, strKart));
      //
      // проверим - есть еще ячейка в строке, если есть значит диапазон
      if( !mat.find() )
        return;
      // есть вторая ячейка, значит диапазон
      int c2 = getExcelColumnNumber(mat.group(1));
      int r2 = Integer.parseInt(mat.group(2));
      // при задании диапазона правая граница д.б. больше левой
      if( c2 < c1 || r2 < r1 ) {
        throw new NumberFormatException("right less that left");
      }
      // заполним диапазон от края до края
      for(int ic = c1; ic <= c2; ic++) {
        for (int jr = r1; jr <= r2; jr++) {
          // добавим ячейку в набор
          this.f_set.add(new yach(jr, ic, prop, strKart));
        }
      }
    } catch (Exception e) {
      System.err.println("?-Error-karta.addStr('" + strKart + "') error conversion: " + e.getMessage());
    }
  }

  /**
   * преобразовать имя столбца Excel в его номер (A-1, B-2, C-3 ...)
   * @param column    строка имени столбца
   * @return номер столбца
   */
  private static int getExcelColumnNumber(String column)
  {
    int result = 0;
    for(int i = 0; i < column.length(); i++) {
      result *= 26;
      result += column.toUpperCase().charAt(i) - 'A' + 1;
    }
    return result;
  }

} // end of class
