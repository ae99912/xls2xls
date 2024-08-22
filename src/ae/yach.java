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
  final int     irow;   // номер строки  1 - 1
  final int     icol;   // номер столбца A - 1, B - 2
  final String  prop;   // свойство ячейки (для контроля переноса значения)
  final String  name;   // имя ячейки (для справки, по программе не нужно)

  /**
   * конструктор ячейки
   * @param iRow  строка
   * @param iCol  колонка
   * @param Prop  свойство ячейки
   * @param Name  имя ячейки (необязательно)
   */
  yach(int iRow, int iCol, String Prop, String... Name)
  {
    if ( iRow < 1 || iCol <1 ) {
      throw new NumberFormatException("number less 1");
    }
    this.irow = iRow;
    this.icol = iCol;
    this.prop = Prop;   // свойство ячейки
    this.name = Name.length > 0? Name[0]: "";   // имя ячейки (строка карты) только для отладки
  }

  /**
   * определить эквивалентность объекта и данного экземпляра (для множества)
   * ячейки равны по своим координатам, свойство не имеет значение (какое будет, то и ладно)
   * @param obj   объект для сравнения
   * @return объекты равны - true, не равны - false
   */
  @Override
  public boolean equals(Object obj) {
    // https://javarush.com/groups/posts/2179-metodih-equals--hashcode-praktika-ispoljhzovanija
    if (getClass() != obj.getClass())
      return false;
    yach yo = (yach)obj;
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
