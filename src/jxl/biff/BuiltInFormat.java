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

package jxl.biff;

import java.util.List;
import jxl.format.Format;

/**
 * The excel string for the various built in formats.  Used to present
 * the cell format information back to the user
 *
 * The difference between this class and the various format object contained
 * in the jxl.write package is that this object contains the Excel strings,
 * not their java equivalents
 */
final class BuiltInFormat implements Format, DisplayFormat
{
  /**
   * The excel format string
   */
  private final String formatString;

  /**
   * The index
   */
  private final int formatIndex;

  /**
   * Constructor
   *
   * @param s the format string
   * @param i the format index
   */
  private BuiltInFormat(String s, int i)
  {
    formatIndex = i;
    formatString = s;
  }

  /**
   * Accesses the excel format string which is applied to the cell
   * Note that this is the string that excel uses, and not the java
   * equivalent
   *
   * @return the cell format string
   */
  public String getFormatString()
  {
    return formatString;
  }

  /**
   * Accessor for the index style of this format
   *
   * @return the index for this format
   */
  public int getFormatIndex()
  {
    return formatIndex;
  }
  /**
   * Accessor to see whether this format has been initialized
   *
   * @return TRUE if initialized, FALSE otherwise
   */
  public boolean isInitialized()
  {
    return true;
  }
  /**
   * Initializes this format with the specified index number
   *
   * @param pos the position of this format record in the workbook
   */
  public void initialize(int pos)
  {
  }

  /**
   * Accessor to determine whether or not this format is built in
   *
   * @return TRUE if this format is a built in format, FALSE otherwise
   */
  public boolean isBuiltIn()
  {
    return true;
  }

  /**
   * Equals method
   *
   * @return TRUE if the two built in formats are equal, FALSE otherwise
   */
  @Override
  public boolean equals(Object o) {
    if (o == this)
      return true;

    if (o instanceof BuiltInFormat bif)
      return (formatIndex == bif.formatIndex);

    return false;
  }

  @Override
  public int hashCode() {
    int hash = 3;
    hash = 59 * hash + this.formatIndex;
    return hash;
  }

  /**
   * The list of built in formats
   */
  public static final List<BuiltInFormat> builtIns = List.of(
    new BuiltInFormat("", 0),
    new BuiltInFormat("0", 1),
    new BuiltInFormat("0.00", 2),
    new BuiltInFormat("#,##0", 3),
    new BuiltInFormat("#,##0.00", 4),
    new BuiltInFormat("($#,##0_);($#,##0)", 5),
    new BuiltInFormat("($#,##0_);[Red]($#,##0)", 6),
    new BuiltInFormat("($#,##0_);[Red]($#,##0)", 7),
    new BuiltInFormat("($#,##0.00_);[Red]($#,##0.00)", 8),
    new BuiltInFormat("0%", 9),
    new BuiltInFormat("0.00%", 10),
    new BuiltInFormat("0.00E+00", 11),
    new BuiltInFormat("# ?/?", 12),
    new BuiltInFormat("# ??/??", 13),
    new BuiltInFormat("dd/mm/yyyy", 14),
    new BuiltInFormat("d-mmm-yy", 15),
    new BuiltInFormat("d-mmm", 16),
    new BuiltInFormat("mmm-yy", 17),
    new BuiltInFormat("h:mm AM/PM", 18),
    new BuiltInFormat("h:mm:ss AM/PM", 19),
    new BuiltInFormat("h:mm", 20),
    new BuiltInFormat("h:mm:ss", 21),
    new BuiltInFormat("m/d/yy h:mm", 22),
    new BuiltInFormat("(#,##0_);(#,##0)", 0x25),
    new BuiltInFormat("(#,##0_);[Red](#,##0)", 0x26),
    new BuiltInFormat("(#,##0.00_);(#,##0.00)", 0x27),
    new BuiltInFormat("(#,##0.00_);[Red](#,##0.00)", 0x28),
    new BuiltInFormat("_(*#,##0_);_(*(#,##0);_(*\"-\"_);(@_)", 0x29),
    new BuiltInFormat("_($*#,##0_);_($*(#,##0);_($*\"-\"_);(@_)", 0x2a),
    new BuiltInFormat("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);(@_)", 0x2b),
    new BuiltInFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);(@_)", 0x2c),
    new BuiltInFormat("mm:ss", 0x2d),
    new BuiltInFormat("[h]mm:ss", 0x2e),
    new BuiltInFormat("mm:ss.0", 0x2f),
    new BuiltInFormat("##0.0E+0", 0x30),
    new BuiltInFormat("@", 0x31)
  );
}

