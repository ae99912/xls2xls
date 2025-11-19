/*
 * Copyright (c) 2023. AE
 * 2025-11-19
 *
 * запись в ячейки Excel строкового значения
 *   str2xls "строка" C53:F56 file.xlsx
 *
 *
 */
package ae;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Set;

public class MainStr {

  public static void main(String[] args) {
    //
    String[] aaa  = new String[3];  // строка область файл
    int     sheet = 0;  // номер листа для обработки
    int     ia    = 0;
    try {
      for(int i = 0; i < args.length; i++) {
        String key = args[i];
        switch (key) {
          case "-?":
            System.out.println(HelpMessage);
            System.exit(1);
            break;

          case "-v":  // отладочный вывод
            R.debug = true;
            break;

          case "-s":  // номер sheet (листа)
            i++;
            sheet = Integer.parseInt(args[i]);  // номер листа
            break;

          default:
            // параметр входной строки
            aaa[ia++] = key;

            break;
        }
      }
      if( ia != 3 )  {
        throw new RuntimeException();
      }
    } catch (Exception e) {
      System.err.println(ErrMessage);
      System.exit(1);
      return;
    }
    //
    //
    // начнем обработку
    //
    R.out("str2xls " + R.Ver);
    R.out("sheet: " + sheet);
    //
    String inpStr  = aaa[0];
    String array   = aaa[1];
    String xlsFile = aaa[2];
    excel eXls     = new excel(xlsFile, sheet);
    // карта ячеек для копирования
    karta k = new karta();
    k.addStr(array, "");
    Set<yach> kar = k.getSet();
    //
    int count = 0; // кол-во записанных значений
    // пройдемся по ячейкам карты
    for(yach ya: kar) {
      int r = ya.irow - 1;    // индекс строки ячейки
      int c = ya.icol - 1;    // индекс столбца ячейки
      //
      Cell cell = eXls.getCell(r, c);   // возьмем ячейку, согласно карте, во входном Excel
      cell.setCellValue(inpStr);
      count++;  // считаем записи
    }
    if(count >0) {
      if( !eXls.write(xlsFile) ) {
        System.err.println("?-Error-don't write: " + xlsFile);
        System.exit(2);
      }
    }
    eXls.close();
  }

  private final static String HelpMessage =
          "str2xls " + R.Ver + "\n" +
                  "Write string to Excel file\n" +
                  "Help about program:\n" +
                  "> str2xls [-v] [-s 0]  string  array  File.xlsx\n" +
                  "-v   отладочный вывод\n" +
                  "-s 0 обрабатываемый лист (sheet) 0, 1 и т.д.";

  private final static String ErrMessage =
          "Неправильный формат командной строки. Смотри -?";


} // end of class
