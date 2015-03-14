package jxl.read.biff;

import java.util.List;

/**
 * created 2015-03-04
 * @author Jan Schlößin
 */
public interface IVerticalPageBreaks {

  public static class ColumnIndex {
    private final int firstColumnFollowingBreak;
    private final int firstRow;
    private final int lastRow;
    
    public ColumnIndex(int firstRowBelowBreak, int firstRow, int lastRow) {
      this.firstColumnFollowingBreak = firstRowBelowBreak;
      this.firstRow = firstRow;
      this.lastRow = lastRow;
    }

    public int getFirstColumnFollowingBreak() {
      return firstColumnFollowingBreak;
    }

    public int getFirstRow() {
      return firstRow;
    }

    public int getLastRow() {
      return lastRow;
    }

    public ColumnIndex withFirstColumnFollowingBreak(int i) {
      return new ColumnIndex(i, firstRow, lastRow);
    }
    
  }

  /**
   * Gets the row breaks
   *
   * @return the row breaks on the current sheet
   */
  List<Integer> getColumnBreaks();
  
}
