/*********************************************************************
*
*      Copyright (C) 2002 Andrew Khan
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
*
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
* Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public
* License along with this library; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
***************************************************************************/

package jxl.read.biff;

import java.util.*;
import jxl.Range;
import jxl.Sheet;
import jxl.biff.IntegerHelper;
import jxl.biff.RecordData;
import jxl.biff.SheetRangeImpl;

/**
 * A merged cells record for a given sheet
 */
public class MergedCellsRecord extends RecordData
{
  /**
   * The ranges of the cells merged on this sheet
   */
  private final List<Range> ranges;

  /**
   * Constructs this object from the raw data
   *
   * @param t the raw data
   * @param s the sheet
   */
  MergedCellsRecord(Record t, Sheet s)
  {
    super(t);

    byte[] data = getRecord().getData();

    int numRanges = IntegerHelper.getInt(data[0], data[1]);
    Range rangesArray[] = new Range[numRanges];
    int pos = 2;
    for (int i = 0; i < numRanges; i++)
    {
      int firstRow = IntegerHelper.getInt(data[pos], data[pos + 1]);
      int lastRow  = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
      int firstCol = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
      int lastCol  = IntegerHelper.getInt(data[pos + 6], data[pos + 7]);

      rangesArray[i] = new SheetRangeImpl(s, firstCol, firstRow,
                                     lastCol, lastRow);

      pos += 8;
    }

    ranges = List.of(rangesArray);
  }

  /**
   * Gets the ranges which have been merged in this sheet
   *
   * @return the ranges of cells which have been merged
   */
  public List<Range> getRanges()
  {
    return ranges;
  }
}







