
/*
 * Copyright (c) 2023. AE
 * 2023-07-11
 *
 * копирование значений ячеек (не пустых) из одного файла Excel 2010 в другой
 * на основе карты переноса:
 *   # пример карты
 *   C53:F56
 *   C53
 *   C99:F99
 *   C110:F110
 */

/*
Modify
 */

package ae;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Set;

public class Main {
  public static void main(String[] args) {
    System.out.println("xls2xls " + R.Ver);
    //
    int ai = 0;
    String[] aaa = new String[3];  // карта входнойфайл выходнойфайл

    for(int i = 0; i < args.length && ai < 3; i++) {
      String key = args[i];

      switch (key) {
        case "-?":
          System.out.println("Help about program:");
          System.out.println("> xls2xls [-v]  Karta.txt  Input.XLSX  Output.XLSX");
          return;
          //break;

        case "-v":  // отладочный вывод
          R.debug = true;
          break;

        default:
          // параметр входной строки
          aaa[ai++] = key;
          break;
      }
    }
    if ( ai < 3 )  {
      System.err.println("Неправильный формат командной строки. Смотри -?");
      return;
    }
    //
    // начнем обработку
    //
    String kartaFile  = aaa[0];
    String inpFile    = aaa[1];
    String outFile    = aaa[2];
    //
    // объекты Excel
    excel eInp = new excel(inpFile, 0);
    excel eOut = new excel(outFile, 0);
    int count = 0;
    // карта ячеек для копирования
    karta k = new karta();
    Set<yach> kar = k.open(kartaFile);
    for(yach ya: kar) {
      int r = ya.irow - 1;
      int c = ya.icol - 1;
      Cell cell = eInp.getCell(r, c);   // возьмем ячейку, согласно карте, во входном Excel
      if(eOut.setCellVal(r, c, cell))   // поместим ячейку в выходной Excel
        count++;  // считаем переносы значений
    }
    //
    if( !eOut.write(outFile) ) {
      System.err.println("?-Error-don't write: " + outFile);
    }
    eInp.close();
    eOut.close();

    System.out.println("Записано ячеек: " + count);
    //
  }

}
