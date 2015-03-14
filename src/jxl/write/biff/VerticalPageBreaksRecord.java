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

import java.util.*;
import jxl.biff.*;

/**
 * Contains the list of explicit horizontal page breaks on the current sheet
 */
class VerticalPageBreaksRecord extends WritableRecordData
{
  /**
   * The row breaks
   */
  private final List<Integer> columnBreaks = new ArrayList<>();
  
  /**
   * Constructor
   * 
   * @param break the row breaks
   */
  public VerticalPageBreaksRecord(List<Integer> breaks)
  {
    super(Type.VERTICALPAGEBREAKS);
    columnBreaks.addAll(breaks);
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

    for (int i = 0; i < columnBreaks.size(); i++)
    {
      IntegerHelper.getTwoBytes(columnBreaks.get(i), data, pos);
      IntegerHelper.getTwoBytes(0xffff, data, pos+4);
      pos += 6;
    }

    return data;
  }
}