package jxl.read.biff;

import java.util.List;

/**
 * created 2015-03-04
 * @author Jan Schlößin
 */
public interface IVerticalPageBreaks {

  /**
   * Gets the row breaks
   *
   * @return the row breaks on the current sheet
   */
  List<Integer> getColumnBreaks();
  
}
