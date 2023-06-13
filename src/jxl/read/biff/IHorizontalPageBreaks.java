package jxl.read.biff;

import java.util.List;

/**
 * created 2015-03-04
 * @author Jan Schlößin
 */
public interface IHorizontalPageBreaks {

  public static class RowIndex {
    private int firstRowBelowBreak;
    private final int firstColumn;
    private final int lastColumn;
    
    public RowIndex(int firstRowBelowBreak, int firstColumn, int lastColumn) {
      this.firstRowBelowBreak = firstRowBelowBreak;
      this.firstColumn = firstColumn;
      this.lastColumn = lastColumn;
    }

    public int getFirstRowBelowBreak() {
      return firstRowBelowBreak;
    }

    public void setFirstRowBelowBreak(int firstRowBelowBreak) {
      this.firstRowBelowBreak = firstRowBelowBreak;
    }

    public int getFirstColumn() {
      return firstColumn;
    }

    public int getLastColumn() {
      return lastColumn;
    }

    public RowIndex withFirstRowBelowBreak(int i) {
      return new RowIndex(i, firstColumn, lastColumn);
    }

    @Override
    public String toString() {
      return "RowBreaks: {" + firstRowBelowBreak + ", first=" + firstColumn + ", last=" + lastColumn + '}';
    }
    
  }
  
  /**
   * Gets the row breaks
   *
   * @return the row breaks on the current sheet
   */
  List<Integer> getRowBreaks();
  
}
