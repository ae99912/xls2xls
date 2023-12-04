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
  09.11.23 указывается номер листа
 */

package ae;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Set;

public class Main {
  public static void main(String[] args) {
    //
    int ia = 0;
    String[] aaa = new String[3];  // карта входнойфайл выходнойфайл
    int   sheet = 0;  // номер листа для обработки

    for(int i = 0; i < args.length; i++) {
      String key = args[i];

      switch (key) {
        case "-?":
          System.out.println(HelpMessage);
          return;
          //break;

        case "-v":  // отладочный вывод
          R.debug = true;
          break;

        case "-s":  // номер sheet (листа)
          i++;
          try {
            sheet = Integer.parseInt(args[i]);  // номер
          } catch (Exception e) {
            System.err.println(ErrMessage);
            return;
          }
          break;

        default:
          // параметр входной строки
          if(ia < aaa.length) {
            aaa[ia++] = key;
          }
          break;
      }
    }
    if ( ia != aaa.length )  {
      System.err.println(ErrMessage);
      return;
    }
    //
    // начнем обработку
    //
    R.out("xls2xls " + R.Ver);
    R.out("sheet: " + sheet);
    //
    String kartaFile  = aaa[0];
    String inpFile    = aaa[1];
    String outFile    = aaa[2];
    //
    // объекты Excel
    excel eInp = new excel(inpFile, sheet);
    excel eOut = new excel(outFile, sheet);
    int count = 0;
    // карта ячеек для копирования
    karta k = new karta();
    Set<yach> kar = k.open(kartaFile);
    boolean only01 = k.isProp("only01");  // записывать только 0 или 1
    //
    for(yach ya: kar) {
      int r = ya.irow - 1;    // индекс строки ячейки
      int c = ya.icol - 1;    // индекс столбца ячейки
      // если берем только ячейки с 0 или 1
      Cell cell = eInp.getCell(r, c);   // возьмем ячейку, согласно карте, во входном Excel
      if(only01) {
        Double d = eInp.getCellNumeric(r,c);
        if(d == null) continue;
        int i01 = d.intValue();
        if(i01!=0 && i01!=1) continue;
      }
      if(eOut.setCellVal(r, c, cell)) {   // поместим ячейку в выходной Excel
        count++;  // считаем переносы значений
      }
    }
    //
    if( !eOut.write(outFile) ) {
      System.err.println("?-Error-don't write: " + outFile);
    }
    eInp.close();
    eOut.close();

    R.out("Записано ячеек: " + count);
    //
  }

  private final static String HelpMessage =
      "xls2xls " + R.Ver + "\n" +
      "Help about program:\n" +
      "> xls2xls [-v] [-s 0]  Karta.txt  Input.XLSX  Output.XLSX\n" +
          "-v   отладочный вывод\n" +
          "-s 0 обрабатываемый лист (sheet) 0, 1 и т.д. всех файлов";

  private final static String ErrMessage =
      "Неправильный формат командной строки. Смотри -?";

}
