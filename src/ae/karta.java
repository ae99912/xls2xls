/*
 * Copyright (c) 2023. AE
 */
/*
  карта переноса значений из входного файла

 */
package ae;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.util.HashSet;
import java.util.Set;

public class karta {
  String    f_filename;   // имя файла карты
  HashSet<yach> f_set;    // набор массива ячеек

  karta()
  {
    f_set = new HashSet<>();
  }

  /**
   * прочитать файл с картой и запомнить ячейки в множестве ячеек cell
   * @param fileName  - имя файла
   * @return набор (множество) ячеек карты
   */
  Set<yach> openSetCells(String fileName)
  {
    try {
      f_filename = fileName;
      f_set.clear();
      File f = new File(fileName);
      FileReader fire = new FileReader(f);
      BufferedReader reader = new BufferedReader(fire);
      String line;
      while( (line = reader.readLine()) != null ) {
        if(line.length() > 1 && line.charAt(0) != '#') {
          yach c = new yach();
          if (c.set(line)) {
            f_set.add(c);
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

  Set<yach> getSetCells()
  {
    return f_set;
  }


} // end of class
