/*********************************************************************
*
*      Copyright (C) 2003 Andrew Khan
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

import java.util.HashMap;
import java.util.Locale;
import java.util.ResourceBundle;

import jxl.common.Logger;

/**
 * A class which contains the function names for the current workbook. The
 * function names can potentially vary from workbook to workbook depending
 * on the locale
 */
public class FunctionNames
{

  /**
   * A hash mapping keyed on the function and returning its locale specific
   * name
   */
  private final HashMap<Function, String> names = new HashMap<>(Function.getFunctions().length);

  /**
   * A hash mapping keyed on the locale specific name and returning the
   * function
   */
  private final HashMap<String, Function> functions = new HashMap<>(Function.getFunctions().length);

  /**
   * Constructor
   *
   * @param l the locale
   */
  public FunctionNames(Locale l)
  {
    ResourceBundle rb = ResourceBundle.getBundle("functions", l);

    // Iterate through all the functions, adding them to the hash maps
    for (Function f : Function.getFunctions()) {
      String propname = f.getPropertyName();
      String n = propname.length() != 0 ? rb.getString(propname) : null;
      if (n != null)
      {
        names.put(f, n);
        functions.put(n, f);
      }
    }
  }

  /**
   * Gets the function for the specified name
   *
   * @param s the string
   * @return  the function
   */
  Function getFunction(String s) {
    return functions.get(s);
  }

  /**
   * Gets the name for the function
   *
   * @param f the function
   * @return  the string
   */
  String getName(Function f) {
    return names.get(f);
  }
}
