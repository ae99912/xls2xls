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

/*
  Заполнение выходного Excel по карте
  A1:G10
  F24

  спец обработка (свойства)
  @only01   дальше заносим только 0 или 1
  @only-01  дальше заносим только 0, 1 или -1
  @onlyint  дальше заносим только целые числа
  @onlynum  дальше заносим только числа действительные или целые
  @all      дальше заносим что угодно
  @blank    дальше ячейки, которые очищаются

  @@regexp - дальше заносим ячейки если они соответствуют regexp
  @=строка - дальше в ячейки заносится "строка"

 */
package ae;

import org.apache.poi.ss.usermodel.*;

import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
    //
    // пройдемся по ячейкам карты
    for(yach ya: kar) {
      int r = ya.irow - 1;    // индекс строки ячейки
      int c = ya.icol - 1;    // индекс столбца ячейки
      //
      Cell cellInp  = eInp.getCell(r, c);   // возьмем ячейку, согласно карте, во входном Excel
      Cell cellOut  = eOut.getCell(r,c);    // выходная ячейка
      //
      boolean isBlank     = false;  // нет очистки содержимого ячейки
      String  strPattern  = null;   // нет паттерна проверки значения
      String  strInsert   = null;   // нет строки вставки
      // свойство данной ячейки
      switch (ya.prop) {
        case "only01":
          strPattern = "[01](\\.0)?";         // только 0 или 1 (целое число завершается .0 , а запятых нет)
          break;

        case "only-01":
          strPattern = "[-01](\\.0)?";        // только - или 0 или 1
          break;

        case "onlyint":
          strPattern = "-?[0-9]+(\\.0)?";     // только целые
          break;

        case "onlynum":
          strPattern = "-?[0-9]+\\.?[0-9]*";  // только числа (целые и действительные)
          break;

        case R.CELL_BLANK:
          isBlank = true;                     // очистка ячейки
          break;

        case R.CELL_ALL:                      // любое значение
          break;

        default:
          if (ya.prop.length() > 1) {
            String s = ya.prop.substring(0, 1);  // буква свойства
            switch (s) {
              case R.PAT_REGEX:               // регулярное выражение
                strPattern = ya.prop.substring(R.PAT_STRING.length());
                break;

              case R.PAT_STRING:              // строка вставки в ячейку
                strInsert = ya.prop.substring(R.PAT_STRING.length());  // строка для вставки
                break;

              default:
                System.err.println("?-warning-неправильное свойство: " + ya.prop);
                continue;
            }
          }
          break;
      }
      // начнем запись в выходную ячейку
      // задана очистка?
      if(isBlank) {
        cellOut.setCellType(Cell.CELL_TYPE_BLANK);
        R.out(cellOut.getAddress() + " - blank");
        count++;  // считаем переносы значений
        continue;
      }
      // определена строка для вставки?
      if(strInsert != null) {
        cellOut.setCellValue(strInsert);
        R.out(cellOut.getAddress() + " - insert: " + strInsert);
        count++;  // вставляем строку в ячейку
        continue;
      }
      // определено ли регулярное выражение для проверки соответствия значения в ячейке?
      if(null != strPattern) {
        // значение string
        String sy = excel.getText(cellInp);
        Pattern pat = Pattern.compile(strPattern);
        Matcher mat = pat.matcher(sy);
        if(!mat.matches())  // сравнивает ВСЮ строку с шаблоном
          continue;           // не соответствует шаблону - пропускаем
      }
      if(excel.copyCell(cellInp, cellOut)) {   // копируем значение ячейки в ячейку выходного Excel
        count++;  // считаем переносы значений
      }
    }
    // выполним вычисления формул, если была запись в ячейки
    if(count > 0) {
      eOut.calculate();
    }
    if( !eOut.write(outFile) ) {
      System.err.println("?-Error-don't write: " + outFile);
    }
    eInp.close();
    eOut.close();
    //
    R.out("Записано ячеек: " + count);
  }

  private final static String HelpMessage =
      "xls2xls " + R.Ver + "\n" +
      "Help about program:\n" +
      "> xls2xls [-v] [-s 0]  Karta.txt  Input.XLSX  Output.XLSX\n" +
          "-v   отладочный вывод\n" +
          "-s 0 обрабатываемый лист (sheet) 0, 1 и т.д. всех файлов\n" +
              "\n" +
              "спец обработка (свойства):\n" +
              "  @only01   дальше заносим только 0 или 1\n" +
              "  @only-01  дальше заносим только 0, 1 или -1\n" +
              "  @onlyint  дальше заносим только целые числа\n" +
              "  @onlynum  дальше заносим только числа действительные или целые\n" +
              "  @all      дальше заносим что угодно\n" +
              "  @blank    дальше ячейки, которые очищаются\n" +
              "\n" +
              "  @@regexp - дальше заносим ячейки если они соответствуют regexp\n" +
              "  @=строка - дальше в ячейки заносится \"строка\"";

  private final static String ErrMessage =
      "Неправильный формат командной строки. Смотри -?";

}  // END OF CLASS MAIN
