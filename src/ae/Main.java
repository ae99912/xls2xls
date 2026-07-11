/*
 * Copyright (c) 2023. AE
 * 11.07.2023
 *
 * Modify:
 * 09.11.23 указывается номер листа
 * 26.12.25 обработка ячейки с датой
 * 27.12.25 вставка строки в ячейку
 * 29.12.25 удалил свойство @only-01, оптимизация чтения входной строки
 * 22.01.26 обработка ячеек строго по порядку в тексте карты
 * 28.03.28 разные листы входного и выходного файла, изменил свойства и добавил @any
 *
*/

/*
 * копирование значений ячеек из одного файла Excel 2010 в другой
 * на основе карты переноса.
 *   # пример карты
 *   C53:F56
 *   D51
 *
 * * спец обработка (свойства):
 *   @int     дальше заносим только целые числа
 *   @num     дальше заносим только действительные или целые числа
 *   @all     дальше заносим что угодно, кроме пустых ячеек
 *   @any     дальше заносим все ячейки, и пустые тоже
 *   @01      дальше заносим только 0 или 1
 *   @012     дальше заносим только 0, 1, 2
 *   @_01     дальше заносим только пустое или 0 или 1
 *   @_012    дальше заносим только пустое или 0, 1, 2
 *   @blank   дальше ячейки, которые очищаются
 *   @@regexp дальше заносим ячейки если они соответствуют regexp
 *   @=строка дальше в ячейки заносится "строка"
 *
 */

package ae;

import org.apache.poi.ss.usermodel.*;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
  public static void main(String[] args) {
    //
    int ia = 0;
    String[] aaa = new String[3];  // карта входной_файл выходной_файл
    int sheetI = 0;  // номер входного листа для обработки
    int sheetO = 0;  // номер выходного листа для обработки

    try {
      for (int i = 0; i < args.length; i++) {
        String arg = args[i];
        switch (arg) {

          case "-?":
            System.out.println(HelpMessage);
            return;

          case "-v":
            R.verbose = true;                   // подробный вывод
            break;

          case "-si":
            i++;
            sheetI = Integer.parseInt(args[i]); // номер входного листа
            break;

          case "-so":
            i++;
            sheetO = Integer.parseInt(args[i]); // номер выходного листа
            break;

          case "-df":
            i++;
            R.DateFormat = args[i];             // формат даты для выходной ячейки
            break;

          default:
            // параметр входной строки
            aaa[ia++] = arg;
            break;
        }
      }
      if(aaa.length != ia) throw new IllegalArgumentException();  // недостаточно аргументов
      //
    } catch (Exception e) {
      System.err.println(ErrMessage);
      System.exit(10);
    }
    //
    // начнем обработку
    //
    R.out("xls2xls " + R.Ver + "   sheet_input:" + sheetI + " sheet_output:" + sheetO);
    //
    String kartaFile = aaa[0];
    String inpFile   = aaa[1];
    String outFile   = aaa[2];
    //
    // объекты Excel
    excel eInp = new excel(inpFile, sheetI);
    excel eOut = new excel(outFile, sheetO);
    int count = 0;
    // карта ячеек для копирования
    karta k = new karta();
    ArrayList<yach> kar = k.open(kartaFile);
    //
    // пройдемся по ячейкам карты
    for(yach ya: kar) {
      int r = ya.irow - 1;    // индекс строки ячейки
      int c = ya.icol - 1;    // индекс столбца ячейки
      //
      Cell cellInp = eInp.getCell(r,c);    // возьмем ячейку, согласно карте, во входном Excel
      Cell cellOut = eOut.getCell(r,c);    // выходная ячейка
      //
      boolean isBlank    = false;   // очистка содержимого ячейки
      boolean isAny      = false;   // любое и пустое значение ячейки
      //
      String  strPattern = null;    // паттерн проверки значения
      String  strInsert  = null;    // строка вставки
      // свойство данной ячейки
      switch (ya.prop) {

        case "int":
          strPattern = "^-?[0-9]+$";            // только целые
          break;

        case "num":
          strPattern = "^-?[0-9]+[.,]?[0-9]*$"; // только числа (целые и действительные)
          break;

        case "01":
          strPattern = "^[0-1]$";               // 0-1
          break;

        case "012":
          strPattern = "^[0-2]$";               // 0-2
          break;

        case "_01":
          strPattern = "(^\\s*$)|(^[0-1]$)";    // пустое или 0-1
          isAny = true;
          break;

        case "_012":
          strPattern = "(^\\s*$)|(^[0-2]$)";    // пустое или 0-2
          isAny = true;
          break;

        case R.CELL_BLANK:
          isBlank = true;                       // очистка ячейки
          break;

        case R.CELL_ALL:                        // любое значение
          break;

        case R.CELL_ANY:                        // любое и пустое значения
          isAny = true;
          break;

        default:
          if (ya.prop.length() > 1) {
            String s = ya.prop.substring(0, 1); // буква свойства
            switch (s) {
              case R.PAT_REGEX:                 // регулярное выражение
                strPattern = ya.prop.substring(R.PAT_REGEX.length());
                isAny = true;   // пусть будут любые (и пустые) значения определенные паттерном
                // https://ru.stackoverflow.com/questions/435544/%D0%9F%D0%BE%D0%B8%D1%81%D0%BA-%D0%BF%D1%83%D1%81%D1%82%D0%BE%D0%B9-%D1%81%D1%82%D1%80%D0%BE%D0%BA%D0%B8-%D0%BF%D0%BE-%D1%80%D0%B5%D0%B3%D1%83%D0%BB%D1%8F%D1%80%D0%BA%D0%B5
                // например, пустая строка: (^\s*$) ^ значит начало строки, \s значит пробельный символ, * значит 0 раз или больше, и $ значит конец строки.
                break;

              case R.PAT_INSTR:                 // строка вставки в ячейку
                strInsert = ya.prop.substring(R.PAT_INSTR.length());  // строка для вставки
                break;

              default:
                System.err.println("?-warning-неправильное свойство: @" + ya.prop);
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
      if(null != strInsert) {
        cellOut.setCellType(Cell.CELL_TYPE_STRING); // явно установить тип вых. ячейки
        cellOut.setCellValue(strInsert);            // вставляем строку в ячейку
        R.out(cellOut.getAddress() + " =\"" + strInsert + "\"");
        count++;
        continue;
      }
      // определено регулярное выражение для проверки соответствия значения в ячейке?
      if(null != strPattern) {
        String  sy  = excel.getText(cellInp);       // текстовое значение ячейки
        Pattern pat = Pattern.compile(strPattern);  // паттерн regex
        Matcher mat = pat.matcher(sy);
        if(!mat.matches())  {   // сравнивает ВСЮ строку с шаблоном
          R.out("? " + cellInp.getAddress() + " (" + sy + ") don't matches \"" + strPattern + "\"");
          continue;   // не соответствует шаблону - пропускаем
        }
      }
      if(excel.copyCell(cellInp, cellOut, isAny)) {  // копируем значение ячейки в ячейку выходного Excel
        count++;      // считаем переносы значений
      }
    }
    R.out("Записано ячеек: " + count);
    // если была запись в ячейки
    if(count > 0) {
      eOut.calculate();             // выполним вычисления формул
      //
      if (!eOut.write(outFile)) {   // запись выходного файла
        System.err.println("?-Error-don't write: " + outFile);
      }
    }
    eInp.close();
    eOut.close();
  }

  private final static String HelpMessage =
      "xls2xls " + R.Ver + "\n" +
      "> xls2xls karta.txt input.xlsx output.xlsx  [-v] [-si 0] [-so 0] [-df " + R.DateFormat + "]\n" +
      "-v           подробный вывод\n" +
      "-si 0        обрабатываемый лист входного файла\n" +
      "-so 0        обрабатываемый лист выходного файла\n" +
      "-df format   формат даты для выходной ячейки (https://clck.ru/3Uey3a)\n" +
      "\n" +
      "Копирует непустые ячейки из входного в выходной лист.\n" +
      "спец обработка (свойства) в последующие ячейки копируем:\n" +
      "  @int       только целые числа\n" +
      "  @num       только числа действительные или целые\n" +
      "  @all       любые непустые значения\n" +
      "  @any       что угодно, в том числе пустые значения\n" +
      "  @01        0 или 1\n" +
      "  @012       0 или 1 или 2\n" +
      "  @_01       пустое или 0 или 1\n" +
      "  @_012      пустое или 0 или 1 или 2\n" +
      "  @@regexp   ячейки если они соответствуют regexp (пустые тоже)\n" +
      "  @blank     пустое значение - ячейки очищаются\n" +
      "  @=строка   строку - в ячейки записывается \"строка\"";

  private final static String ErrMessage =
      "Неправильный формат командной строки. Смотри -?";

}  // END OF CLASS MAIN
