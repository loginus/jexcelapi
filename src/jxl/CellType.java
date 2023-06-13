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

package jxl;

/**
 * An enumeration type listing the available content types for a cell
 */
public enum CellType {

  // An empty cell can still contain formatting information and comments
  EMPTY("Empty"),
  LABEL("Label"),
  NUMBER("Number"),
  BOOLEAN("Boolean"),
  ERROR("Error"),
  NUMBER_FORMULA("Numerical Formula"),
  DATE_FORMULA("Date Formula"),
  STRING_FORMULA("String Formula"),
  BOOLEAN_FORMULA("Boolean Formula"),
  FORMULA_ERROR("Formula Error"),
  DATE("Date");

  private final String description;

  private CellType(String desc) {
    description = desc;
  }

  @Override
  public String toString()
  {
    return description;
  }

}


