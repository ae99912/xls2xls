/*
 * Copyright (c) 2023. AE
 */

/*
  обработка данных
 */
package ae;

/*
  Работа над Excel
 */

import org.apache.poi.ss.usermodel.Cell;

import java.util.Set;

public class work {

  int up(String kartaFile, String inpFile, String outFile)
  {
    excel einp = new excel(inpFile, 0);
    excel eout = new excel(outFile, 0);
    int count = 0;
    String str;
    Double dbl;
    int r, c;
    karta k = new karta();
    Set<yach> kar = k.openSetCells(kartaFile);
    for(yach ya: kar) {
      r = ya.irow - 1;
      c = ya.icol - 1;
      Cell cc = einp.getCell(r, c); //
      int type = cc.getCellType();  // тип ячейки
      switch (type) {
        // строка
        case Cell.CELL_TYPE_STRING:
          str = cc.getStringCellValue();
          if( str != null && str.length() > 0) {
            eout.setCellVal(r, c, str);
            count++;
          }
          break;

        // число
        case Cell.CELL_TYPE_NUMERIC:
          dbl = cc.getNumericCellValue();
          if( dbl != null ) {
            eout.setCellVal(r, c, dbl);
            count++;
          }
          break;
      }
      //
    } // end for

    //
    eout.write(outFile);
    einp.close();
    eout.close();

    return count;
  }

} // end of class
