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

package jxl.write.biff;

import java.util.List;
import jxl.biff.IntegerHelper;
import jxl.biff.Type;
import jxl.biff.WritableRecordData;
import jxl.read.biff.VerticalPageBreaksRecord.ColumnIndex;

/**
 * Contains the list of explicit horizontal page breaks on the current sheet
 */
class VerticalPageBreaksRecord extends WritableRecordData
{
  /**
   * The row breaks
   */
  private List<ColumnIndex> columnBreaks;
  
  /**
   * Constructor
   * 
   * @param break the row breaks
   */
  public VerticalPageBreaksRecord(List<ColumnIndex> breaks)
  {
    super(Type.VERTICALPAGEBREAKS);

    columnBreaks = breaks;
  }

  /**
   * Gets the binary data to write to the output file
   * 
   * @return the binary data
   */
  public byte[] getData()
  {
    byte[] data = new byte[columnBreaks.size() * 6 + 2];

    // The number of breaks on the list
    IntegerHelper.getTwoBytes(columnBreaks.size(), data, 0);
    int pos = 2;

    for (ColumnIndex cb : columnBreaks) {
      IntegerHelper.getTwoBytes(cb.getFirstColumnFollowingBreak(), data, pos);
      IntegerHelper.getTwoBytes(cb.getFirstRow(), data, pos+2);
      IntegerHelper.getTwoBytes(cb.getLastRow(), data, pos+4);
      pos += 6;
    }

    return data;
  }
}


