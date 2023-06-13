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

package jxl.write.biff;

import java.io.IOException;
import java.util.*;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Cell;
import jxl.CellType;
import jxl.Range;
import jxl.WorkbookSettings;
import jxl.biff.SheetRangeImpl;
import jxl.write.Blank;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

/**
 * Contains all the merged cells, and the necessary logic for checking
 * for intersections and for handling very large amounts of merging
 */
class MergedCells
{
  /**
   * The logger
   */
  private static final Logger logger = Logger.getLogger(MergedCells.class);

  /**
   * The list of merged cells
   */
  private List<SheetRangeImpl> ranges = new ArrayList<>();

  /**
   * The sheet containing the cells
   */
  private final WritableSheet sheet;

  /**
   * The maximum number of ranges per sheet
   */
  private static final int maxRangesPerSheet = 1020;

  /**
   * Constructor
   */
  public MergedCells(WritableSheet ws)
  {
    sheet = ws;
  }

  /**
   * Adds the range to the list of merged cells.  Does no checking
   * at this stage
   *
   * @param range the range to add
   */
  void add(SheetRangeImpl r)
  {
    ranges.add(r);
  }

  /**
   * Used to adjust the merged cells following a row insertion
   */
  void insertRow(int row)
  {
    // Adjust any merged cells
    for (SheetRangeImpl sr : ranges)
      sr.insertRow(row);
  }

  /**
   * Used to adjust the merged cells following a column insertion
   */
  void insertColumn(int col)
  {
    for (SheetRangeImpl sr : ranges)
      sr.insertColumn(col);
  }

  /**
   * Used to adjust the merged cells following a column removal
   */
  void removeColumn(int col)
  {
    Iterator<SheetRangeImpl> i = ranges.iterator();
    while (i.hasNext())
    {
      SheetRangeImpl sr = i.next();
      if (sr.getTopLeft().getColumn() == col &&
          sr.getBottomRight().getColumn() == col)
      {
        // The column with the merged cells on has been removed, so get
        // rid of it from the list
        i.remove();
      }
      else
      {
        sr.removeColumn(col);
      }
    }
  }

  /**
   * Used to adjust the merged cells following a row removal
   */
  void removeRow(int row)
  {
    Iterator<SheetRangeImpl> i = ranges.iterator();
    while (i.hasNext())
    {
      SheetRangeImpl sr = i.next();
      if (sr.getTopLeft().getRow()   == row &&
          sr.getBottomRight().getRow() == row)
      {
        // The row with the merged cells on has been removed, so get
        // rid of it from the list
        i.remove();
      }
      else
      {
        sr.removeRow(row);
      }
    }
  }

  /**
   * Gets the cells which have been merged on this sheet
   *
   * @return an unmodifiable view of range objects
   */
  List<SheetRangeImpl> getMergedCells()
  {
    return Collections.unmodifiableList(ranges);
  }

  /**
   * Unmerges the specified cells.  The Range passed in should be one that
   * has been previously returned as a result of the getMergedCells method
   *
   * @param r the range of cells to unmerge
   */
  void unmergeCells(Range r)
  {
    int index = ranges.indexOf(r);

    if (index != -1)
    {
      ranges.remove(index);
    }
  }

  /**
   * Called prior to writing out in order to check for intersections
   */
  private void checkIntersections()
  {
    ArrayList<SheetRangeImpl> newcells = new ArrayList<>(ranges.size());

    for (SheetRangeImpl r : ranges) {
      // Check that the range doesn't intersect with any existing range
      Iterator<SheetRangeImpl> i = newcells.iterator();
      boolean intersects = false;
      while (i.hasNext() && !intersects)
      {
        SheetRangeImpl range = i.next();

        if (range.intersects(r))
        {
          logger.warn("Could not merge cells " + r +
                  " as they clash with an existing set of merged cells.");

          intersects = true;
        }
      }

      if (!intersects)
      {
        newcells.add(r);
      }
    }

    ranges = newcells;
  }

  /**
   * Checks the cell ranges for intersections, or if the merged cells
   * contains more than one item of data
   */
  private void checkRanges()
  {
    try
    {
      // Check all the ranges to make sure they only contain one entry
      for (SheetRangeImpl range : ranges) {

        // Get the cell in the top left
        Cell tl = range.getTopLeft();
        Cell br = range.getBottomRight();
        boolean found = false;

        for (int c = tl.getColumn(); c <= br.getColumn(); c++)
        {
          for (int r = tl.getRow(); r <= br.getRow(); r++)
          {
            Cell cell = sheet.getCell(c, r);
            if (cell.getType() != CellType.EMPTY)
            {
              if (!found)
              {
                found = true;
              }
              else
              {
                logger.warn("Range " + range +
                        " contains more than one data cell.  " +
                                "Setting the other cells to blank.");
                Blank b = new Blank(c, r);
                sheet.addCell(b);
              }
            }
          }
        }
      }
    }
    catch (WriteException e)
    {
      // This should already have been checked - bomb out
      Assert.verify(false);
    }
  }

  void write(File outputFile) throws IOException
  {
    if (ranges.isEmpty())
      return;

    WorkbookSettings ws =
      ( (WritableSheetImpl) sheet).getWorkbookSettings();

    if (!ws.getMergedCellCheckingDisabled())
    {
      checkIntersections();
      checkRanges();
    }

    // If they will all fit into one record, then create a single
    // record, write them and get out
    if (ranges.size() < maxRangesPerSheet)
    {
      MergedCellsRecord mcr = new MergedCellsRecord(ranges);
      outputFile.write(mcr);
      return;
    }

    int numRecordsRequired = ranges.size() / maxRangesPerSheet + 1;
    int pos = 0;

    for (int i = 0 ; i < numRecordsRequired ; i++)
    {
      int numranges = Math.min(maxRangesPerSheet, ranges.size() - pos);

      ArrayList<SheetRangeImpl> cells = new ArrayList<>(numranges);
      for (int j = 0 ; j < numranges ; j++)
      {
        cells.add(ranges.get(pos+j));
      }

      MergedCellsRecord mcr = new MergedCellsRecord(cells);
      outputFile.write(mcr);

      pos += numranges;
    }
  }
}
