
package jxl;

/**
 * created 30.05.2020
 * @author jan
 */
public class CellCoordinate implements Comparable<CellCoordinate> {

  private final int column;
  private final int row;

  public CellCoordinate(int column, int row) {
    this.column = column;
    this.row = row;
  }

  public int getColumn() {
    return column;
  }

  public int getRow() {
    return row;
  }

  @Override
  public int hashCode() {
    int hash = 7;
    hash = 37 * hash + this.column;
    hash = 37 * hash + this.row;
    return hash;
  }

  @Override
  public boolean equals(Object obj) {
    if (this == obj)
      return true;
    if (obj == null)
      return false;
    if (getClass() != obj.getClass())
      return false;
    return compareTo((CellCoordinate) obj) == 0;
  }

  @Override
  public int compareTo(CellCoordinate that) {
    int rowDiff = Integer.compare(this.row, that.row);
    if (rowDiff != 0)
      return rowDiff;
    else
      return Integer.compare(this.column, that.column);
  }

}
