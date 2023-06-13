package jxl;

import java.util.*;
import jxl.write.*;

/**
 * This class describes the locatation of a cell whithin a workbook.
 * It realizes the value object pattern.
 *
 * created 2020-01-05
 * @author jan
 */
public class CellLocation {

  private final WritableSheet sheet;
  private final int column, row;

  public CellLocation(WritableSheet sheet, int column, int row) {
    this.sheet = sheet;
    this.column = column;
    this.row = row;
  }

  public WritableSheet getSheet() {
    return sheet;
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
    hash = 37 * hash + Objects.hashCode(this.sheet);
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
    final CellLocation other = (CellLocation) obj;
    if (this.column != other.column)
      return false;
    if (this.row != other.row)
      return false;
    return Objects.equals(this.sheet, other.sheet);
  }

}
