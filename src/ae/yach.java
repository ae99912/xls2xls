/*
 * Copyright (c) 2023. AE
 */

/*
 * класс Ячейка,
 * храним номер строки и столбца, сформированных на основе названия ячейки
 *
 */
package ae;

class yach {
  int irow;       // номер строки  1 - 1
  int icol;       // номер столбца A - 1, B - 2
  String  name;   // имя ячейки (для справки, по программе не нужно)

  /**
   * установить значения
   * @param iRow  строка
   * @param iCol  колонка
   * @param Name  имя
   */
  yach(int iRow, int iCol,  String Name)
  {
    set(iRow, iCol, Name);
  }

  void set(int iRow,int iCol,  String Name) throws NumberFormatException
  {
    if ( iRow < 1 || iCol <1 ) {
      throw new NumberFormatException("number less 1");
    }
    this.irow = iRow;
    this.icol = iCol;
    if(R.debug) this.name = Name;   // имя ячейки (строка карты) только для отладки
  }

  /**
   * определить эквивалентность объекта и данного экземпляра (для множества)
   * @param o   объект для сравнения
   * @return объекты равны - true, не равны - false
   */
  @Override
  public boolean equals(Object o) {
    // https://javarush.com/groups/posts/2179-metodih-equals--hashcode-praktika-ispoljhzovanija
    if (getClass() != o.getClass())
      return false;
    yach yo = (yach)o;
    return this.icol == yo.icol  && this.irow == yo.irow;
  }

  /**
   * выдать хэш код для объекта (для множества)
   * @return хэш-код
   */
  @Override
  public  int hashCode() {
    // https://javarush.com/groups/posts/2179-metodih-equals--hashcode-praktika-ispoljhzovanija
    String s = this.icol + "," + this.irow;
    return s.hashCode();
  }

} // end of class
