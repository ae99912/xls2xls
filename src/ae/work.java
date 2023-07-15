/*
 * Copyright (c) 2023. AE
 */

/*
  обработка данных
 */

/*
  Работа над Excel
 */
package ae;

import org.apache.poi.ss.usermodel.Cell;
import java.util.Set;

public class work {

  int up(String kartaFile, String inpFile, String outFile)
  {
    excel eInp = new excel(inpFile, 0);
    excel eOut = new excel(outFile, 0);
    int count = 0;
    // карта ячеек для копирования
    karta k = new karta();
    Set<yach> kar = k.openSetYach(kartaFile);
    for(yach ya: kar) {
      int r = ya.irow - 1;
      int c = ya.icol - 1;
      Cell cc = eInp.getCell(r, c); //
      if(eOut.setCellVal(r, c, cc))
        count++;
    }
    //
    if( !eOut.write(outFile) ) {
      System.err.println("?-Error-don't write: " + outFile);
    }
    eInp.close();
    eOut.close();
    //
    return count;
  }

} // end of class
