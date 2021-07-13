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

package jxl.biff.formula;

import java.util.*;

/**
 * An enumeration detailing the Excel parsed tokens
 * A particular token may be associated with more than one token code
 */
class Token
{
  /**
   * The array of values which apply to this token
   */
  private final byte[] values;

  /**
   * All available tokens, keyed on value
   */
  private static Map<Byte, Token> tokens = new HashMap<>(100);

  /**
   * Constructor
   * Sets the token value and adds this token to the array of all token
   *
   * @param reference the biff code for the token
   */
  private Token(int reference)
  {
    values = new byte[] {(byte) reference};

    tokens.put((byte) reference, this);
    System.out.println(tokens.size());
  }

  /**
   * Constructor
   * Sets the token value and adds this token to the array of all token
   *
   * @param v the biff code for the token
   */
  private Token(int reference, int value, int array)
  {
    this.values = new byte[] {(byte) reference, (byte) value, (byte) array};

    tokens.put((byte) reference, this);
    tokens.put((byte) value, this);
    tokens.put((byte) array, this);
    System.out.println(tokens.size());
  }

  /**
   * Gets the reference token code for the specified token.  This is always
   * the first on the list
   *
   * @return the token code.  This is the first item in the array
   */
  public byte getReferenceCode()
  {
    return values[0];
  }

  /**
   * Gets the value token code for the specified token.  This is always
   * the second item on the list
   *
   * @return the token code
   */
  public byte getValueCode()
  {
    return values.length > 0 ? values[1] : values[0];
  }

  public byte getArrayCode() {
    return values.length > 1 ? values[2] : getValueCode();
  }

  /**
   * Gets the type object from its integer value
   */
  public static Token getToken(byte v)
  {
    Token t = tokens.get(v);

    return t != null ? t : UNKNOWN;
  }

  // Binary Operators
  public static final Token ADD           = new Token(0x3);
  public static final Token SUBTRACT      = new Token(0x4);
  public static final Token MULTIPLY      = new Token(0x5);
  public static final Token DIVIDE        = new Token(0x6);
  public static final Token POWER         = new Token(0x7);
  public static final Token CONCAT        = new Token(0x8);
  public static final Token LESS_THAN     = new Token(0x9);
  public static final Token LESS_EQUAL    = new Token(0xa);
  public static final Token EQUAL         = new Token(0xb);
  public static final Token GREATER_EQUAL = new Token(0xc);
  public static final Token GREATER_THAN  = new Token(0xd);
  public static final Token NOT_EQUAL     = new Token(0xe);
  public static final Token INTERSECTION  = new Token(0xf);
  public static final Token UNION         = new Token(0x10);
  public static final Token RANGE         = new Token(0x11);

  // Unary Operators
  public static final Token UNARY_PLUS   = new Token(0x12);
  public static final Token UNARY_MINUS  = new Token(0x13);
  public static final Token PERCENT      = new Token(0x14);

  // Operands
  public static final Token MISSING_ARG = new Token(0x16);
  public static final Token STRING      = new Token(0x17);
  public static final Token ERR         = new Token(0x1c);
  public static final Token BOOL        = new Token(0x1d);
  public static final Token INTEGER     = new Token(0x1e);
  public static final Token DOUBLE      = new Token(0x1f);
  public static final Token ARRAY       = new Token(0x20, 0x40, 0x60);
  public static final Token NAMED_RANGE = new Token(0x23, 0x43, 0x63); //need 0x23 for data validation references
  public static final Token REF         = new Token(0x24, 0x44, 0x64);
  public static final Token AREA        = new Token(0x25, 0x45, 0x65);
  public static final Token MEM_AREA    = new Token(0x26, 0x46, 0x66);
  public static final Token MEM_ERR     = new Token(0x27, 0x47, 0x67);
  public static final Token REFERR      = new Token(0x2a, 0x4a, 0x6a);
  public static final Token AREA_R      = new Token(0x2b, 0x4b, 0x6b); // AREA_ERR
  public static final Token REFV        = new Token(0x2c, 0x4c, 0x6c); // REF_N
  public static final Token AREAV       = new Token(0x2d, 0x4d, 0x6d); // AREA_N
  public static final Token NAME        = new Token(0x39, 0x59, 0x79);
  public static final Token REF3D       = new Token(0x3a, 0x5a, 0x7a);
  public static final Token AREA3D      = new Token(0x3b, 0x5b, 0x7b);
  public static final Token REF_ERR_3D  = new Token(0x3c, 0x5c, 0x7c);
  public static final Token AREA_ERR_3D = new Token(0x3d, 0x5d, 0x7d);

  // Function operators
  public static final Token FUNCTION       = new Token(0x21, 0x41, 0x61);
  public static final Token FUNCTIONVARARG = new Token(0x22, 0x42, 0x62);
  public static final Token MACROCOMMAND   = new Token(0x38, 0x58, 0x78); // BIFF2,BIFF3

  // Control
  public static final Token EXP         = new Token(0x1);
  public static final Token TBL         = new Token(0x2);
  public static final Token PARENTHESIS = new Token(0x15);
  public static final Token NLR         = new Token(0x18); // BIFF8, extended parsed thing
  public static final Token ATTRIBUTE   = new Token(0x19);
  public static final Token SHEET       = new Token(0x1A); // BIFF2-4, deleted
  public static final Token END_SHEET   = new Token(0x1B); // BIFF2-4, deleted
  public static final Token MEM_NO_MEM  = new Token(0x28, 0x48, 0x68);
  public static final Token MEM_FUNC    = new Token(0x29, 0x49, 0x69);
  public static final Token MEM_AREA_V  = new Token(0x2e, 0x4e, 0x6e); // MEM_AREA_N
  public static final Token MEM_NO_MEM_V= new Token(0x2f, 0x4f, 0x6f); // MEM_AREA_N

  // Unknown token
  public static final Token UNKNOWN = new Token(0xffff);
}