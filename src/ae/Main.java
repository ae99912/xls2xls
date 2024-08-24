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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
  final static String Name_regex = "regex";    // имя свойства "регулярное выражение"
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
    //
    // пройдемся по ячейкам карты
    for(yach ya: kar) {
      int r = ya.irow - 1;    // индекс строки ячейки
      int c = ya.icol - 1;    // индекс столбца ячейки
      //
      Cell cell = eInp.getCell(r, c);   // возьмем ячейку, согласно карте, во входном Excel
      //
      // строка паттерна регулярного выражения
      String strPattern;
      // свойство данной ячейки
      switch(ya.prop) {
        case "only01":
          strPattern = "[01](\\.0)?";        // только 0 или 1 (целое число завершается .0 , а запятых нет)
          break;

        case "only-01":
          strPattern = "[-01](\\.0)?";       // только - или 0 или 1
          break;

        case "onlyint":
          strPattern = "-?[0-9]+(\\.0)?";    // только целые
          break;

        case "onlynum":
          strPattern = "-?[0-9]+\\.?[0-9]*";   // только числа (целые и действительные)
          break;

        default:
          if(ya.prop.startsWith(Name_regex)) {    // это свойство "регулярное выражение"?
            // строка после имени свойства - само регулярное выражение
            strPattern = ya.prop.substring(Name_regex.length());
          } else {
            strPattern = null;  // нет свойства - паттерн пустой
          }
          break;
      }
      // определено ли регулярное выражение для проверки соответствия значения в ячейке?
      if(strPattern != null) {
        String sy = excel.getCellStrValue(cell);    // значение ячейки
        Pattern pat = Pattern.compile(strPattern);
        Matcher mat = pat.matcher(sy);
        if( !mat.matches() )  // сравнивает ВСЮ строку с шаблоном
          continue;           // не соответствует шаблону - пропускаем
      }
      if(eOut.setCellTo(cell, r, c)) {   // поместим ячейку в выходной Excel (строка, колонка)
        count++;  // считаем переносы значений
      }
    }
    //
    if( !eOut.write(outFile) ) {
      System.err.println("?-Error-don't write: " + outFile);
    }
    eInp.close();
    eOut.close();
    //
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

}  // END OF CLASS MAIN
