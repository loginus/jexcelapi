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

/**
 * An enumeration class which contains the biff types
 */
public enum Type
{
  BOF(0x809),
  EOF(0x0a),
  BOUNDSHEET(0x85),
  SUPBOOK(0x1ae),
  EXTERNSHEET(0x17),
  DIMENSION(0x200),
  BLANK(0x201),
  MULBLANK(0xbe),
  ROW(0x208),
  NOTE(0x1c),
  TXO(0x1b6),
  RK (0x7e),
  RK2 (0x27e),
  MULRK (0xbd),
  INDEX(0x20b),
  DBCELL(0xd7),
  SST(0xfc),
  COLINFO(0x7d),
  EXTSST(0xff),
  CONTINUE(0x3c),
  LABEL(0x204),
  RSTRING(0xd6),
  LABELSST(0xfd),
  NUMBER(0x203),
  NAME(0x18),
  TABID(0x13d),
  ARRAY(0x221),
  STRING(0x207),
  FORMULA(0x406),
  FORMULA2(0x6),
  SHAREDFORMULA(0x4bc),
  FORMAT(0x41e),
  XF(0xe0),
  BOOLERR(0x205),
  INTERFACEHDR(0xe1),
  SAVERECALC(0x5f),
  INTERFACEEND(0xe2),
  XCT(0x59),
  CRN(0x5a),
  DEFCOLWIDTH(0x55),
  DEFAULTROWHEIGHT(0x225),
  WRITEACCESS(0x5c),
  WSBOOL(0x81),
  CODEPAGE(0x42),
  DSF(0x161),
  FNGROUPCOUNT(0x9c),
  FILTERMODE(0x9b),
  AUTOFILTERINFO(0x9d),
  AUTOFILTER(0x9e),
  COUNTRY(0x8c),
  PROTECT(0x12),
  SCENPROTECT(0xdd),
  OBJPROTECT(0x63),
  PRINTHEADERS(0x2a),
  HEADER(0x14),
  FOOTER(0x15),
  HCENTER(0x83),
  VCENTER(0x84),
  FILEPASS(0x2f),
  SETUP(0xa1),
  PRINTGRIDLINES(0x2b),
  GRIDSET(0x82),
  GUTS(0x80),
  WINDOWPROTECT(0x19),
  PROT4REV(0x1af),
  PROT4REVPASS(0x1bc),
  PASSWORD(0x13),
  REFRESHALL(0x1b7),
  WINDOW1(0x3d),
  WINDOW2(0x23e),
  BACKUP(0x40),
  HIDEOBJ(0x8d),
  NINETEENFOUR(0x22),
  PRECISION(0xe),
  BOOKBOOL(0xda),
  FONT(0x31),
  MMS(0xc1),
  CALCMODE(0x0d),
  CALCCOUNT(0x0c),
  REFMODE(0x0f),
  TEMPLATE(0x60),
  OBJPROJ(0xd3),
  DELTA(0x10),
  MERGEDCELLS(0xe5),
  ITERATION(0x11),
  STYLE(0x293),
  USESELFS(0x160),
  VERTICALPAGEBREAKS(0x1a),
  HORIZONTALPAGEBREAKS(0x1b),
  SELECTION(0x1d),
  HLINK(0x1b8),
  OBJ(0x5d),
  MSODRAWING(0xec),
  MSODRAWINGGROUP(0xeb),
  LEFTMARGIN(0x26),
  RIGHTMARGIN(0x27),
  TOPMARGIN(0x28),
  BOTTOMMARGIN(0x29),
  EXTERNNAME(0x23),
  PALETTE(0x92),
  PLS(0x4d),
  SCL(0xa0),
  PANE(0x41),
  WEIRD1(0xef),
  SORT(0x90),
  CONDFMT(0x1b0),
  CF(0x1b1),
  DV(0x1be),
  DVAL(0x1b2),
  BUTTONPROPERTYSET(0x1ba),
  EXCEL9FILE(0x1c0),

  // Chart types
  FONTX(0x1026),
  IFMT(0x104e),
  FBI(0x1060),
  ALRUNS(0x1050),
  SERIES(0x1003),
  SERIESLIST(0x1016),
  SBASEREF(0x1048),
  UNKNOWN(0xffff),

  // Pivot stuff

  // Unknown types
  U1C0(0x1c0),
  U1C1(0x1c1);

  /**
   * The biff value for this type
   */
  public final int value;

  /**
   * Constructor
   *
   * @param v the biff code for the type
   */
  private Type(int v)
  {
    value = v;
  }

  /**
   * Gets the type object from its integer value
   * @param v the internal code
   * @return the type
   */
  public static Type getType(int v)
  {
    for (Type t : values())
      if (t.value == v)
        return t;

    return UNKNOWN;
  }

}