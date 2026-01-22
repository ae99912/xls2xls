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
  22.01.26 вместо множества ввел список, чтобы были последовательные значения ячеек

 */
package ae;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class karta {
  // паттерн для имени ячейки
  final private Pattern cell_pattern = Pattern.compile("([A-Z]+)([0-9]+)");  // паттерн для имени ячейки A12, B3 ...

  ArrayList<yach> f_yac = new ArrayList<>();      // набор ячеек

  /**
   * открыть и прочитать файл с картой и запомнить ячейки в множестве ячеек yach
   * @param fileName  имя файла
   * @return набор (множество) ячеек карты
   */
  ArrayList<yach> open(String fileName)
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
    } catch (Exception e) {
      System.err.println("?-Error-karta.open('" + fileName + "') " + e.getMessage());
      return null;
    }
    return f_yac;
  }

  /**
   * добавить в множество ячеек ячейки из строки карты как имя отдельной ячейки или диапазона ячеек
   * @param strKart   строка карты переноса
   */
  public void addStr(String strKart, String prop)
  {
    // буквы в верхний регистр и уберем все пробелы
    String sss = strKart.toUpperCase().replaceAll("\\s", "");
    if( sss.isEmpty() )
      return;
    //
    try {
      Matcher mat = cell_pattern.matcher(sss);
      if( !mat.find() ) {
        throw new NumberFormatException("not found cell name");
      }
      int c1 = getExcelColumnNumber(mat.group(1));
      int r1 = Integer.parseInt(mat.group(2));
      // проверим - есть еще ячейка в строке, если есть значит диапазон
      int c2 = c1;
      int r2 = r1;
      if( mat.find() ) {
        // есть вторая ячейка, значит диапазон
        c2 = getExcelColumnNumber(mat.group(1));
        r2 = Integer.parseInt(mat.group(2));
      }
      // при задании диапазона правая граница д.б. больше левой
      if( c2 < c1 || r2 < r1 ) {
        throw new NumberFormatException("right less that left");
      }
      // заполним диапазон от края до края
      for(int ic = c1; ic <= c2; ic++) {
        for (int jr = r1; jr <= r2; jr++) {
          // добавим ячейку в набор
          String s = getExcelColumnName(ic) + jr + " (" + strKart + ")"; // строка нужна для коммента и для отладки
          this.f_yac.add(new yach(jr, ic, prop, s));
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
  static int getExcelColumnNumber(String column)
  {
    int result = 0;
    for(int i = 0; i < column.length(); i++) {
      result *= 26;
      result += column.toUpperCase().charAt(i) - 'A' + 1;
    }
    return result;
  }

  /**
   * преобразовать номер столбца Excel в его имя (1-A, 2-B, 3-C, ...)
   * @param numcol    номер столбца
   * @return имя столбца
   */
  static String getExcelColumnName(int numcol)
  {
    final StringBuilder sb = new StringBuilder();
    int num = numcol - 1;
    while(num >= 0) {
      int numChar = (num % 26) + 65;
      sb.append((char)numChar);
      num = (num / 26) - 1;
    }
    return sb.reverse().toString();
  }

} // end of class
