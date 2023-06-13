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
enum Token {

    // Binary Operators
    ADD          (0x3),
    SUBTRACT     (0x4),
    MULTIPLY     (0x5),
    DIVIDE       (0x6),
    POWER        (0x7),
    CONCAT       (0x8),
    LESS_THAN    (0x9),
    LESS_EQUAL   (0xa),
    EQUAL        (0xb),
    GREATER_EQUAL(0xc),
    GREATER_THAN (0xd),
    NOT_EQUAL    (0xe),
    INTERSECTION (0xf),
    UNION        (0x10),
    RANGE        (0x11),

    // Unary Operators
    UNARY_PLUS   (0x12),
    UNARY_MINUS  (0x13),
    PERCENT      (0x14),

    // Constant Operands
    MISSING_ARG (0x16),
    STRING      (0x17),
    ERR         (0x1c),
    BOOL        (0x1d),
    INTEGER     (0x1e),
    DOUBLE      (0x1f),
    ARRAY       (0x20, 0x40, 0x60), // the reference class 0x20 never appears in an excel formula

    // Operands
    NAME        (0x23, 0x43, 0x63), //need 0x23 for data validation references
    REF         (0x24, 0x44, 0x64),
    AREA        (0x25, 0x45, 0x65),
    MEM_AREA    (0x26, 0x46, 0x66),
    MEM_ERR     (0x27, 0x47, 0x67),
    REFERR      (0x2a, 0x4a, 0x6a),
    AREA_ERR    (0x2b, 0x4b, 0x6b),
    REF_N       (0x2c, 0x4c, 0x6c),
    AREA_N      (0x2d, 0x4d, 0x6d),
    NAME_X      (0x39, 0x59, 0x79), // local name of explicit sheet or external name
    REF3D       (0x3a, 0x5a, 0x7a),
    AREA3D      (0x3b, 0x5b, 0x7b),
    REF_ERR_3D  (0x3c, 0x5c, 0x7c),
    AREA_ERR_3D (0x3d, 0x5d, 0x7d),

    // Function operators
    FUNCTION       (0x21, 0x41, 0x61),
    FUNCTIONVARARG (0x22, 0x42, 0x62),
    MACROCOMMAND   (0x38, 0x58, 0x78), // BIFF2,BIFF3

    // Control
    EXP         (0x1),
    TBL         (0x2),
    PARENTHESIS (0x15),
    NLR         (0x18), // BIFF8, extended parsed thing
    ATTRIBUTE   (0x19),
    SHEET       (0x1A), // BIFF2-4, deleted
    END_SHEET   (0x1B), // BIFF2-4, deleted
    MEM_NO_MEM  (0x28, 0x48, 0x68),
    MEM_FUNC    (0x29, 0x49, 0x69),
    MEM_AREA_N  (0x2e, 0x4e, 0x6e),
    MEM_NO_MEM_N(0x2f, 0x4f, 0x6f),

    // Unknown token
    UNKNOWN (0);

  /**
   * The array of values which apply to this token
   */
  private final byte[] values;

  /**
   * Constructor
   * Sets the token value and adds this token to the array of all token
   *
   * @param reference the biff code for the token
   */
  private Token(int reference) {
    values = new byte[] {(byte) reference};
  }

  /**
   * Constructor
   * Sets the token value and adds this token to the array of all token
   *
   * @param reference the biff code for the token
   * @param value the biff code for the token
   * @param array the biff code for the token
   */
  private Token(int reference, int value, int array) {
    this.values = new byte[] {(byte) reference, (byte) value, (byte) array};
  }

  /**
   * Gets the reference token code for the specified token.  This is always
   * the first on the list
   *
   * @return the token code. This is the first item in the array
   *
   * TODO: used in DATA_VALIDATION = getReferenceCode = ! alternateCode
   */
  public byte getReferenceCode() {
    return values[0];
  }

  /**
   * Gets the value token code for the specified token.  This is always
   * the second item on the list
   *
   * @return the token code
   *
   * TODO: used in DEFAULT = getValueCode = alternateCode
   */
  public byte getValueCode() {
    return values.length > 1 ? values[1] : values[0];
  }

  /**
   * Gets the array token code for the specified token.  This is always
   * the third item on the list
   *
   * @return the token code
   */
  public byte getArrayCode() {
    return values.length > 2 ? values[2] : getValueCode();
  }

  /**
   * All available tokens, keyed on value
   */
  private static final Map<Byte, Token> TOKENS = new HashMap<>(100);

  static {
    Arrays.stream(Token.values())
            .peek(t -> TOKENS.put(t.getReferenceCode(), t))
            .peek(t -> TOKENS.put(t.getValueCode(), t))
            .forEach(t -> TOKENS.put(t.getArrayCode(), t));
  }

  /**
   * Gets the type object from its byte value
   */
  public static Token getToken(byte v) {
    return TOKENS.getOrDefault(v, UNKNOWN);
  }

}