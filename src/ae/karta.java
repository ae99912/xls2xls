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

public class karta {
  HashSet<yach> f_set;    // набор множества ячеек

  karta()
  {
    f_set = new HashSet<>();
  }

  /**
   * прочитать файл с картой и запомнить ячейки в множестве ячеек yach
   * @param fileName  имя файла
   * @return набор (множество) ячеек карты
   */
  Set<yach> openSetYach(String fileName)
  {
    try {
      f_set.clear();
      BufferedReader rdr = new BufferedReader(new FileReader(fileName));
      String str;
      while( (str = rdr.readLine()) != null ) {
        if( str.length() > 1 && str.charAt(0) != '#' ) {
          yach ya = new yach();
          Set<yach> setYach = ya.set(str);
          if(setYach != null)
            f_set.addAll(setYach);
//          if (ya.set(str)) {
//            f_set.add(ya);
//          }
        }
      }
      //
    } catch (Exception e) {
      System.err.println("?-Error-karta.open('" + fileName + "') " + e.getMessage());
      return null;
    }
    return f_set;
  }

} // end of class
