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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class excel {
  //
  Workbook      f_wbk       = null;   // workbook Excel рабочая книга
  Sheet         f_sheet     = null;   // sheet Excel рабочий лист

  excel()
  {}

  excel(String fileName, int numSheet)
  {
    if( !open(fileName, numSheet) ) {
      System.err.println("?-Error-excel('" + fileName + "'," + numSheet + ") don't open worksheet");
    }
  }

  /**
   * Открыть рабочую книгу Excel и лист
   * @param fileName  имя файла
   * @param numSheet  индекс листа
   * @return true - открыто, false - не открыто
   */
  boolean open(String fileName, int numSheet)
  {
    if( openWorkbook(fileName) ) {
      return openSheet(numSheet);
    }
    return false;
  }

  /**
   * открыть рабочую книгу Excel
   * @param fileName  имя файла
   * @return true - открыто, false - не открыто
   */
  boolean openWorkbook(String fileName)
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
   * выдать текстовое содержание ячейки
   * @param cell ячейка
   * @return  строка содержимого
   */
  static String getText(Cell cell)
  {
    DataFormatter formatter = new DataFormatter();
    String str = formatter.formatCellValue(cell);
    return str;
  }

  /**
   * Копировать входную ячейку в выходную, взяв результат формулы
   * @param inpCell входная ячейка
   * @param outCell выходная ячейка
   * @return true если копирование выполнено
   */
  static boolean copyCell(Cell inpCell, Cell outCell)
  {
    try {
      int type = inpCell.getCellType();
      // значение string
      switch (type) {
        case Cell.CELL_TYPE_FORMULA:
          int typeo = inpCell.getCachedFormulaResultType();
          inpCell.setCellType(typeo);
          return copyCell(inpCell, outCell);

        case Cell.CELL_TYPE_BLANK:
          outCell.setCellType(Cell.CELL_TYPE_BLANK);
          break;

        case Cell.CELL_TYPE_STRING:
          outCell.setCellValue(inpCell.getStringCellValue());
          break;

        case Cell.CELL_TYPE_BOOLEAN:
          outCell.setCellValue(inpCell.getBooleanCellValue());
          break;

        case Cell.CELL_TYPE_NUMERIC:
          if (DateUtil.isCellDateFormatted(inpCell)) {
            String str = getText(inpCell);
            // проверим на дату
            String rx = "([0-2]{1,2})\\/([0-9]{1,2})\\/([0-9]{2})";
            Pattern pat = Pattern.compile(rx);
            Matcher mat = pat.matcher(str);
            if(mat.find()) {
              // найдена дата
              int m,d,y;
              m = Integer.parseInt(mat.group(1));
              d = Integer.parseInt(mat.group(2));
              y = Integer.parseInt(mat.group(3));
              str = String.format("%02d.%02d.%04d", d,m,2000+y);
            }
            outCell.setCellValue(str);  // запишем строковое значение
          } else {
            outCell.setCellValue(inpCell.getNumericCellValue());
          }
          break;

        default:
          return false;
      }
    } catch (Exception e) {
      System.err.println("?-error-copyCell(): " + e.getMessage());
    }
    return true;
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