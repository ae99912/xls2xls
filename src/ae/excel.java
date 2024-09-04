/*
 * Copyright (c) 2023. AE
 */

/*
 Книга Excel, на основе файла на диске
 */
package ae;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class excel {
  //
  Workbook      f_wbk       = null;   // workbook Excel рабочая книга
  Sheet         f_sheet     = null;   // sheet Excel рабочий лист

  excel(String fileName, int numSheet)
  {
    if( open(fileName) ) {
      if( !openSheet(numSheet) ) {
        System.err.println("?-Error-excel('" + fileName + "'," + numSheet + ") don't open worksheet");
      }
    }
  }

  /**
   * открыть рабочую книгу Excel
   * @param fileName  имя файла
   * @return true - открыто, false - не открыто
   */
  boolean open(String fileName)
  {
    if(f_wbk != null) {
      System.err.println("?-Warning-open('" + fileName + "') workbook already open");
      return false;
    }
    try {
      f_sheet = null;
      // File tmpFile = copyToTmp(fileName);
      FileInputStream inp = new FileInputStream(fileName);
      f_wbk = new XSSFWorkbook(inp); // прочитать файл с Excel 2010
      inp.close();
    } catch (Exception e) {
      System.err.println("?-Error-open('" + fileName + "')  " + e.getMessage());
      return false;
    }
    return true;
  }

  /**
   * открыть лист в Excel
   * @param numSheet  индекс листа
   * @return результат true - лист открыт, false - не открыт
   */
  boolean openSheet(int numSheet)
  {
    if(f_wbk == null) {
      System.err.println("?-Warning-openSheet('" + numSheet + "') don't open worksheet");
      return false;
    }
    try {
      f_sheet = f_wbk.getSheetAt(numSheet); //Access the worksheet, so that we can update / modify it.
    } catch (Exception e) {
      System.err.println("?-Error-openSheet('" + numSheet + ")  " + e.getMessage());
      f_sheet = null;
      return false;
    }
    return true;
  }
  /**
   * закрыть объект
   */
  void close()
  {
    if(f_wbk != null) {
      try {
        f_wbk.close();
      } catch (Exception e) {
        System.err.println("?-Error-close() " + e.getMessage());
      }
      f_wbk   = null;
      f_sheet = null;
    }
  }

  /**
   * записать рабочую книгу Excel в выходной файл
   * @param fileName  имя выходного файла
   * @return результат записи (true - записан, false - нет)
   */
  boolean write(String fileName)
  {
    if(f_wbk == null) {
      System.err.println("?-Error-write('" + fileName + "') don't open Excel");
      return false;
    }
    try {
      // Write the output to a file
      // при отладке (при прерывании исполнения) выходной файл портится, поэтому
      // сначала запишем Excel во временный файл
      File tempFile = File.createTempFile("x2x",".tmp");
      tempFile.deleteOnExit();  // удалить при завершении приложения
      FileOutputStream ftmpout = new FileOutputStream(tempFile);
      f_wbk.write(ftmpout);
      ftmpout.close();
      // если после записи во временный файл его длина больше 1 кБ, то запишем в выходной файл
      if(tempFile.length() > 1024) {
        File f = new File(fileName);
        // копирование файла
        // https://javadevblog.com/kak-skopirovat-fajl-v-java-4-sposoba-primery-i-kod.html
        Files.copy(tempFile.toPath(), f.toPath(), REPLACE_EXISTING);
        return true;
      }
    } catch (Exception e) {
      System.err.println("?-Error-write('" + fileName +"') " + e.getMessage());
      return false;
    }
    return false;
  }

  /**
   * записать в ячейку таблицы значение по указанной ячейке
   * @param cell  ячейка, откуда берется значение
   * @param irow  строка ячейки, куда помещаем значение входной ячейки
   * @param icol  колонка ячейки
   * @return  результат записи - было записано значение или нет
   */
  boolean setCellTo(Cell cell, int irow, int icol)
  {
    try {
      Cell c = getCell(irow, icol);
      int type = cell.getCellType();  // тип ячейки
      if(type == Cell.CELL_TYPE_FORMULA) {
        // если формула, то поставим ее значение
        type = cell.getCachedFormulaResultType();
        R.out("?-Warning-setCellTo(" + getCellStrValue(cell) + ", " + irow + "," + icol + ") formula: " + cell.getCellFormula());
      }
      boolean r = setCellTypeContent(cell, type, c);
      R.out("setCellTo(" + getCellStrValue(cell) + ", " + irow + "," + icol + ")");
      return r;
    } catch (Exception e) {
      System.err.println("?-Error-setCellTo(" + getCellStrValue(cell) + ", " + irow + "," + icol + ")-error set value. " + e.getMessage());
      return false;
    }
  }

  /**
   * Задает значение выходной ячейки по входной ячейке и ее типу.
   * Используется для записи копий значений или значений формулы
   * https://stackoverflow.com/questions/62305485/how-to-get-function-expression-apache-poi
   * @param cell    входная ячейка (может быть с формулой)
   * @param type    тип вычисленного значения для выходной ячейки
   * @param cellOut выходная ячейка
   * @return строка значения
   */
  private static boolean setCellTypeContent(Cell cell, int type, Cell cellOut) {
    switch (type) {
      case Cell.CELL_TYPE_STRING:
        //System.out.println("String: " + cell.getRichStringCellValue().getString());
        cellOut.setCellValue(cell.getRichStringCellValue().getString());
        break;

      case Cell.CELL_TYPE_NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
          //System.out.println("Date: " + cell.getDateCellValue());
          cellOut.setCellValue(cell.getDateCellValue());
        } else {
          //System.out.println("Number: " + cell.getNumericCellValue());
          cellOut.setCellValue(cell.getNumericCellValue());
        }
        break;

      case Cell.CELL_TYPE_BOOLEAN:
        //System.out.println("Boolean: " + cell.getBooleanCellValue());
        cellOut.setCellValue(cell.getBooleanCellValue());
        break;

      case Cell.CELL_TYPE_FORMULA:
        //System.out.print("Formula result is ");
        cellOut.setCellValue("recursive formula");
        break;

      case Cell.CELL_TYPE_BLANK:
        // System.out.println("Blank cell.");
        // не изменяем выходную ячейку - cellOut.setCellValue("");
        return false;

      default:
        System.err.println("?-Warning-setCellTypeContent(...," + type + ",...). This should not have happened.");
        return false;

    }
    return true;
  }

  /**
   * Получить ячейку в строке в заданной колонке
   * @param irow   строка
   * @param icol   колонка
   * @return  ячейка, null - ошибка (совсем нет ячейки)
   */
  Cell getCell(int irow, int icol)
  {
    if(f_sheet == null) {
      System.err.println("?-Error-getCell(" + irow + "," + icol + ")  don't open worksheet.");
      return null;
    }
    Cell c;
    try {
      Row row = f_sheet.getRow(irow);
      if(row == null) {
        row = f_sheet.createRow(irow);
      }
      c = row.getCell(icol);  // Access the cell
      if(c == null) {
        c = row.createCell(icol); // создадим ячейку
      }
    } catch (Exception e) {
      System.err.println("?-Error-getCell(" + irow + "," + icol  + ") " + e.getMessage());
      return null;
    }
    return c;
  }

  /**
   * Получить из ячейки число
   * @param irow  строка
   * @param icol  колонка
   * @return действительное число, если ячейка есть и в ней число, либо null
   */
  Double getCellNumeric(int irow, int icol)
  {
    Double dbl = null;
    Cell c = getCell(irow, icol);
    if(c != null) {
      if (c.getCellType() == Cell.CELL_TYPE_NUMERIC) {  // numeric 0
        dbl = c.getNumericCellValue();
      }
    }
    return dbl;
  }

  String  getCellString(int irow, int icol)
  {
    String str = null;
    Cell c = getCell(irow, icol);
    if(c != null) {
      str = getCellStrValue(c);
    }
    return str;
  }

  public static String  getCellStrValue(Cell cell)
  {
    if(cell == null)
      return "null";
    return getCellStrValue(cell.getCellType(), cell);
  }

  public static String  getCellStrValue(int type, Cell cell)
  {
    if(cell == null)
      return "null";
    switch (type) { // тип ячейки
      // строка
      case Cell.CELL_TYPE_STRING:
        return cell.getStringCellValue();

      // число
      case Cell.CELL_TYPE_NUMERIC:
        return "" + cell.getNumericCellValue();
      // break;

      case Cell.CELL_TYPE_BOOLEAN:
        return "" + cell.getBooleanCellValue();

      // формула
      case Cell.CELL_TYPE_FORMULA:
        //return "=" + cell.getCellFormula();
        // если формула, то поставим ее значение
        int typeres = cell.getCachedFormulaResultType();
        return getCellStrValue(typeres, cell);

      // бланк
      case Cell.CELL_TYPE_BLANK:
        return "<blank>";

      // ошибка
      case Cell.CELL_TYPE_ERROR:
        return "<error>";

    }
    return "<...>";
  }


} // end of class

  /*

   // выполним принудительно перерасчет всех формул в рабочей книге

  void calculate()
  {
    if(f_wbk == null) {
      System.err.println("?-Error-excel.calculate() don't open Excel");
    }
    // После заполнения ячеек формулы не пересчитываются, поэтому выполним принудительно
    // перерасчет всех формул на листе
    // http://poi.apache.org/spreadsheet/eval.html#Re-calculating+all+formulas+in+a+Workbook
    FormulaEvaluator evaluator = f_wbk.getCreationHelper().createFormulaEvaluator();
    for(Sheet sheet: f_wbk) { for(Row row: sheet) { for(Cell c: row) { if (c.getCellType() == Cell.CELL_TYPE_FORMULA) { evaluator.evaluateFormulaCell(c); } } } }
  }
  */