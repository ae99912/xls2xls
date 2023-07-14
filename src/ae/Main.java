
/*
 * Copyright (c) 2023. AE
 * 2023-07-11
 *
 * копирование значений ячеек (не пустых) из одного файла Excel в другой
 * на основе карты переноса
 */

/*
Modify
 */

package ae;

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
          System.out.println("Help about programm");
          System.out.println("xls2xls karta.txt Input.XLSX Output.XLSX");
          return;
          //break;

        default:
          // параметр входной строки
          aaa[ai++] = key;
          break;
      }
    }
    if (ai < 3)  {
      System.err.println("Неправильный формат командной строки. Смотри -?");
      return;
    }
    //
    work w = new work();
    int cnt;
    // начнем обработку
    cnt = w.up(aaa[0], aaa[1], aaa[2]);
    //
    System.out.println("Записано ячеек: " + cnt);
  }

}
