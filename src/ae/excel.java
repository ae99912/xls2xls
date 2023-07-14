/*
 * Copyright (c) 2023. AE
 */
package ae;
/*
 Книга Excel, на основе файла на диске
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class excel {
  //
  Workbook      f_wbk       = null;
  Sheet         f_sheet     = null;
  String        f_filename  = null;

  excel(String fileName, int numSheet)
  {
    open(fileName, numSheet);
  }

  /**
   * открыть файл Excel
   * @param fileName  - имя файла
   * @param numSheet  - номер листа, с которым работаем
   * @return true - открыто, false - не открыто
   */
  boolean open(String fileName, int numSheet)
  {
    if(f_wbk != null) {
      System.err.println("?-Error-excel.open('" + fileName + "') yet open Excel");
      return false;
    }
    try {
      // File tmpFile = copyToTmp(fileName);
      FileInputStream inp = new FileInputStream(fileName);
      f_wbk = new XSSFWorkbook(inp); // прочитать файл с Excel 2010
      inp.close();
      f_sheet = f_wbk.getSheetAt(numSheet); //Access the worksheet, so that we can update / modify it.
    } catch (Exception e) {
      System.err.println("?-Error-excel.open('" + fileName + "', " +numSheet +")  " + e.getMessage());
      f_sheet = null;
      f_wbk   = null;
      return false;
    }
    return true;
  }

//  /**
//   * копировать входной файл во временный файл
//   * @param fileName  - имя входного файла
//   * @return  временный файл
//   * @throws Exception
//   */
//  File copyToTmp(String fileName) throws Exception
//  {
//    File tempFile = File.createTempFile("x2x",".tmp");
//    tempFile.deleteOnExit();  // удалить при завершении приложения
//    File source = new File(fileName);
//    Files.copy(source.toPath(), tempFile.toPath(), REPLACE_EXISTING);
//    return tempFile;
//  }

  void close()
  {
    if(f_wbk != null) {
      try {
        f_wbk.close();
      } catch (Exception e) {
        System.err.println("?-Error-excel.close()  " + e.getMessage());
      }
      f_wbk = null;
      f_sheet = null;
    }
  }

  boolean write(String fileName)
  {
    if(f_wbk == null) {
      System.err.println("?-Error-excel.write('" + fileName + "') don't open Excel");
      return false;
    }
    try {
      // Write the output to a file
      // при отладке (при прерывании исполнения) выходной файл портится, поэтому
      // сначала запишем Excel во временный файл
      File tempFile = File.createTempFile("x2x",".tmp");
      tempFile.deleteOnExit();  // удалить при завершении приложения
      FileOutputStream fto = new FileOutputStream(tempFile);
      f_wbk.write(fto);
      fto.close();
      // если после записи во временный файл его длина больше 1 кБ, то запишем в выходной файл
      if ( tempFile.length() > 1024 ) {
        File f = new File(fileName);
        // копирование файла
        // https://javadevblog.com/kak-skopirovat-fajl-v-java-4-sposoba-primery-i-kod.html
        Files.copy(tempFile.toPath(), f.toPath(), REPLACE_EXISTING);
      }
    } catch (Exception e) {
      System.err.println("?-Error-excel.write('" + fileName +"') " + e.getMessage());
      return false;
    }
    return true;
  }

  /**
   * выполним принудительно перерасчет всех формул в рабочей книге
   */
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

  /**
   * Установить числовое значение ячейки в заданной строке таблицы
   * @param irow    строка
   * @param icol    номер колонки
   * @param val     устанавливаемое значения (numeric)
   * @return      1 - значение установлено, 0 - не установлено
   */
  boolean setCellVal(int irow, int icol, Double val)
  {
    try {
      Cell c = getCell(irow, icol);
      if(c == null)
        return false;
      c.setCellValue(val);  // Access the cell
    } catch (Exception e) {
      System.err.println("?-Warning-setCellVal(" + irow + "," + icol + ", " + val + ")-error set value. " + e.getMessage());
      return false;
    }
    return true;
  }

  /**
   * Установить строковое значение ячейки в заданной строке таблицы
   * @param irow    строка
   * @param icol    колонка
   * @param val     устанавливаемое значения (String)
   * @return      1 - значение установлено, 0 - не установлено
   */
  boolean setCellVal(int irow, int icol, String val)
  {
    try {
      Cell c = getCell(irow, icol);
      if(c == null)
        return false;
      c.setCellValue(val);  // Access the cell
    } catch (Exception e) {
      System.err.println("?-Warning-setCellVal(" + irow + "," + icol + ", " + val + ")-error set value. " + e.getMessage());
      return false;
    }
    return true;
  }

  /**
   * Получить ячейки в строке в заданной колонке
   * @param irow   строка
   * @param icol   колонка
   * @return  ячейка, null - нет ячейки
   */
  Cell getCell(int irow, int icol)
  {
    if(f_sheet == null) {
      System.err.println("?-Error-excel.getCell(" + irow + "," + icol + ")  don't open Excel.");
      return null;
    }
    Cell c;
    try {
      Row row = f_sheet.getRow(irow);
      c = row.getCell(icol);  // Access the cell
      if (c == null) {
        c = row.createCell(icol); // создадим ячейку
      }
    } catch (Exception e) {
      System.err.println("?-Error-getCell(" + irow + "," + icol  + ") " + e.getMessage());
      return null;
    }
    return c;
  }

//  int getCellType(int irow, int icol)
//  {
//    String str = null;
//    int typeCell = Cell.CELL_TYPE_ERROR;
//    Cell c = getCell(irow, icol);
//    if(c != null) {
//      typeCell = c.getCellType();
//    }
//    return typeCell;
//  }
//
//  String getCellStr(int irow, int icol)
//  {
//    String str = null;
//    Cell c = getCell(irow, icol);
//    if(c != null) {
//      if (c.getCellType() == Cell.CELL_TYPE_STRING) {   // string 1
//        str = c.getStringCellValue();
//      }
//    }
//    return str;
//  }
//
//  Double getCellNumeric(int irow, int icol)
//  {
//    Double dbl = null;
//    Cell c = getCell(irow, icol);
//    if(c != null) {
//      if (c.getCellType() == Cell.CELL_TYPE_NUMERIC) {  // numeric 0
//        dbl = c.getNumericCellValue();
//      }
//    }
//    return dbl;
//  }

} // end of class
