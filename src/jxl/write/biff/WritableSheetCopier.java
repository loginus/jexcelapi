/*********************************************************************
*
*      Copyright (C) 2006 Andrew Khan
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

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Range;
import jxl.WorkbookSettings;
import jxl.biff.CellReferenceHelper;
import jxl.biff.DataValidation;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.NumFormatRecordsException;
import jxl.biff.SheetRangeImpl;
import jxl.biff.WorkspaceInformationRecord;
import jxl.biff.XFRecord;
import jxl.biff.drawing.DrawingGroupObject;
import jxl.format.CellFormat;
import jxl.biff.formula.FormulaException;
import jxl.write.Blank;
import jxl.write.Boolean;
import jxl.write.DateTime;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableHyperlink;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * A transient utility object used to copy sheets.   This
 * functionality has been farmed out to a different class
 * in order to reduce the bloat of the WritableSheetImpl
 */
class WritableSheetCopier
{
  private static Logger logger = Logger.getLogger(SheetCopier.class);

  private final WritableSheetImpl fromSheet;
  private final WritableSheetImpl toSheet;
  private final WorkbookSettings workbookSettings;

  // Objects used by the sheet
  private TreeSet<ColumnInfoRecord> fromColumnFormats;
  private TreeSet<ColumnInfoRecord> toColumnFormats;
  private MergedCells fromMergedCells;
  private MergedCells toMergedCells;
  private RowRecord[] fromRows;
  private HorizontalPageBreaksRecord fromRowBreaks;
  private VerticalPageBreaksRecord fromColumnBreaks;
  private HorizontalPageBreaksRecord toRowBreaks;
  private VerticalPageBreaksRecord toColumnBreaks;
  private DataValidation fromDataValidation;
  private DataValidation toDataValidation;
  private SheetWriter sheetWriter;
  private ArrayList<DrawingGroupObject> fromDrawings;
  private ArrayList<DrawingGroupObject> toDrawings;
  private ArrayList<WritableImage> toImages;
  private WorkspaceInformationRecord fromWorkspaceOptions;
  private PLSRecord fromPLSRecord;
  private PLSRecord toPLSRecord;
  private ButtonPropertySetRecord fromButtonPropertySet;
  private ButtonPropertySetRecord toButtonPropertySet;
  private ArrayList<WritableHyperlink> fromHyperlinks;
  private ArrayList<WritableHyperlink> toHyperlinks;
  private ArrayList<WritableCell> validatedCells;
  private int numRows;
  private int maxRowOutlineLevel;
  private int maxColumnOutlineLevel;


  private boolean chartOnly;
  private FormattingRecords formatRecords;



  // Objects used to maintain state during the copy process
  private HashMap<Integer, WritableCellFormat> xfRecords;
  private HashMap<Integer, Integer> fonts;
  private HashMap<Integer, Integer> formats;

  public WritableSheetCopier(WritableSheet f, WritableSheet t)
  {
    fromSheet = (WritableSheetImpl) f;
    toSheet = (WritableSheetImpl) t;
    workbookSettings = toSheet.getWorkbook().getSettings();
    chartOnly = false;
  }

  void setColumnFormats(TreeSet<ColumnInfoRecord> fcf, TreeSet<ColumnInfoRecord> tcf)
  {
    fromColumnFormats = fcf;
    toColumnFormats = tcf;
  }

  void setMergedCells(MergedCells fmc, MergedCells tmc)
  {
    fromMergedCells = fmc;
    toMergedCells = tmc;
  }

  void setRows(RowRecord[] r)
  {
    fromRows = r;
  }

  void setValidatedCells(ArrayList<WritableCell> vc)
  {
    validatedCells = vc;
  }

  void setRowBreaks(HorizontalPageBreaksRecord frb, HorizontalPageBreaksRecord trb)
  {
    fromRowBreaks = frb;
    toRowBreaks = trb;
  }

  void setColumnBreaks(VerticalPageBreaksRecord fcb, VerticalPageBreaksRecord tcb)
  {
    fromColumnBreaks = fcb;
    toColumnBreaks = tcb;
  }

  void setDrawings(ArrayList<DrawingGroupObject> fd, ArrayList<DrawingGroupObject> td, ArrayList<WritableImage> ti)
  {
    fromDrawings = fd;
    toDrawings = td;
    toImages = ti;
  }

  void setHyperlinks(ArrayList<WritableHyperlink> fh, ArrayList<WritableHyperlink> th)
  {
    fromHyperlinks = fh;
    toHyperlinks = th;
  }

  void setWorkspaceOptions(WorkspaceInformationRecord wir)
  {
    fromWorkspaceOptions = wir;
  }

  void setDataValidation(DataValidation dv)
  {
    fromDataValidation = dv;
  }

  void setPLSRecord(PLSRecord plsr)
  {
    fromPLSRecord = plsr;
  }

  void setButtonPropertySetRecord(ButtonPropertySetRecord bpsr)
  {
    fromButtonPropertySet = bpsr;
  }

  void setSheetWriter(SheetWriter sw)
  {
    sheetWriter = sw;
  }


  DataValidation getDataValidation()
  {
    return toDataValidation;
  }

  PLSRecord getPLSRecord()
  {
    return toPLSRecord;
  }

  boolean isChartOnly()
  {
    return chartOnly;
  }

  ButtonPropertySetRecord getButtonPropertySet()
  {
    return toButtonPropertySet;
  }

  /**
   * Copies a sheet from a read-only version to the writable version.
   * Performs shallow copies
   */
  public void copySheet()
  {
    shallowCopyCells();
    // Copy the column formats

    for (ColumnInfoRecord toCopy : fromColumnFormats)
      toColumnFormats.add(new ColumnInfoRecord(toCopy));

    // Copy the merged cells
    Range[] merged = fromMergedCells.getMergedCells();

    for (Range m : merged)
      toMergedCells.add(new SheetRangeImpl((SheetRangeImpl) m, toSheet));

    try
    {
      for (int i = 0; i < fromRows.length ; i++)
      {
        RowRecord row = fromRows[i];

        if (row != null &&
            (!row.isDefaultHeight() ||
             row.isCollapsed()))
        {
          RowRecord newRow = toSheet.getRowRecord(i);
          newRow.setRowDetails(row.getRowHeight(),
                               row.matchesDefaultFontHeight(),
                               row.isCollapsed(),
                               row.getOutlineLevel(),
                               row.getGroupStart(),
                               row.getStyle());
        }
      }
    }
    catch (RowsExceededException e)
    {
      // Handle the rows exceeded exception - this cannot occur since
      // the sheet we are copying from will have a valid number of rows
      Assert.verify(false);
    }

    // Copy the horizontal page breaks
    toRowBreaks.setRowBreaks(fromRowBreaks);

    // Copy the vertical page breaks
    toColumnBreaks.setColumnBreaks(fromColumnBreaks);

    // Copy the data validations
    if (fromDataValidation != null)
    {
      toDataValidation = new DataValidation
        (fromDataValidation,
         toSheet.getWorkbook(),
         toSheet.getWorkbook(),
         toSheet.getWorkbook().getSettings());
    }

    // Copy the charts
    sheetWriter.setCharts(fromSheet.getCharts());

    // Copy the drawings
    for (DrawingGroupObject o : fromDrawings)
      if (o instanceof jxl.biff.drawing.Drawing drawing) {
        WritableImage wi = new WritableImage
                  (drawing,
                          toSheet.getWorkbook().getDrawingGroup());
        toDrawings.add(wi);
        toImages.add(wi);
      }

    // Not necessary to copy the comments, as they will be handled by
    // the deep copy of the individual cells

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromWorkspaceOptions);

    // Copy the environment specific print record
    if (fromPLSRecord != null)
    {
      toPLSRecord = new PLSRecord(fromPLSRecord);
    }

    // Copy the button property set
    if (fromButtonPropertySet != null)
    {
      toButtonPropertySet = new ButtonPropertySetRecord(fromButtonPropertySet);
    }

    // Copy the hyperlinks
    for (WritableHyperlink toCopy : fromHyperlinks)
      toHyperlinks.add(new WritableHyperlink(toCopy, toSheet));

  }

  /**
   * Performs a shallow copy of the specified cell
   */
  private WritableCell shallowCopyCell(Cell cell) {
    return switch (cell.getType()) {
      case LABEL -> new Label((LabelCell) cell);
      case NUMBER -> new Number((NumberCell) cell);
      case DATE -> new DateTime((DateCell) cell);
      case BOOLEAN -> new Boolean((BooleanCell) cell);
      case NUMBER_FORMULA -> new ReadNumberFormulaRecord((FormulaData) cell);
      case STRING_FORMULA -> new ReadStringFormulaRecord((FormulaData) cell);
      case BOOLEAN_FORMULA -> new ReadBooleanFormulaRecord((FormulaData) cell);
      case DATE_FORMULA -> new ReadDateFormulaRecord((FormulaData) cell);
      case FORMULA_ERROR -> new ReadErrorFormulaRecord((FormulaData) cell);
      case EMPTY -> (cell.getCellFormat() != null)
            // It is a blank cell, rather than an empty cell, so
            // it may have formatting information, so
            // it must be copied
            ? new Blank(cell)
            : null;
      case ERROR -> null;
    };
  }

  /**
   * Performs a deep copy of the specified cell, handling the cell format
   *
   * @param cell the cell to copy
   */
  private WritableCell deepCopyCell(Cell cell)
  {
    WritableCell c = shallowCopyCell(cell);

    if (c == null)
    {
      return c;
    }

    if (c instanceof ReadFormulaRecord rfr)
    {
      boolean crossSheetReference = !rfr.handleImportedCellReferences
        (fromSheet.getWorkbook(),
         fromSheet.getWorkbook(),
         workbookSettings);

      if (crossSheetReference)
      {
        try
        {
        logger.warn("Formula " + rfr.getFormula() +
                    " in cell " +
                    CellReferenceHelper.getCellReference(cell.getColumn(),
                                                         cell.getRow()) +
                    " cannot be imported because it references another " +
                    " sheet from the source workbook");
        }
        catch (FormulaException e)
        {
          logger.warn("Formula  in cell " +
                      CellReferenceHelper.getCellReference(cell.getColumn(),
                                                           cell.getRow()) +
                      " cannot be imported:  " + e.getMessage());
        }

        // Create a new error formula and add it instead
        c = new Formula(cell.getColumn(), cell.getRow(), "\"ERROR\"");
      }
    }

    // Copy the cell format
    CellFormat cf = c.getCellFormat();
    int index = ( (XFRecord) cf).getXFIndex();
    WritableCellFormat wcf = xfRecords.get(index);

    if (wcf == null)
      wcf = copyCellFormat(cf);

    c.setCellFormat(wcf);

    return c;
  }

  /**
   * Perform a shallow copy of the cells from the specified sheet into this one
   */
  void shallowCopyCells()
  {
    // Copy the cells
    for (int i = 0;  i < fromSheet.getRows(); i++) {
      Cell[] row = fromSheet.getRow(i);

      for (Cell cell : row) {
        WritableCell c = shallowCopyCell(cell);

        // Encase the calls to addCell in a try-catch block
        // These should not generate any errors, because we are
        // copying from an existing spreadsheet.  In the event of
        // errors, catch the exception and then bomb out with an
        // assertion
        try
        {
          if (c != null)
          {
            toSheet.addCell(c);

            // Cell.setCellFeatures short circuits when the cell is copied,
            // so make sure the copy logic handles the validated cells
            if (c.getCellFeatures() != null &
                    c.getCellFeatures().hasDataValidation())
            {
              validatedCells.add(c);
            }
          }
        }
        catch (WriteException e)
        {
          Assert.verify(false);
        }
      }
    }
    numRows = toSheet.getRows();
  }

  /**
   * Returns an initialized copy of the cell format
   *
   * @param cf the cell format to copy
   * @return a deep copy of the cell format
   */
  private WritableCellFormat copyCellFormat(CellFormat cf)
  {
    try
    {
      // just do a deep copy of the cell format for now.  This will create
      // a copy of the format and font also - in the future this may
      // need to be sorted out
      XFRecord xfr = (XFRecord) cf;
      WritableCellFormat f = new WritableCellFormat(xfr);
      formatRecords.addStyle(f);

      // Maintain the local list of formats
      int xfIndex = xfr.getXFIndex();
      xfRecords.put(xfIndex, f);

      int fontIndex = xfr.getFontIndex();
      fonts.put(fontIndex, f.getFontIndex());

      int formatIndex = xfr.getFormatRecord();
      formats.put(formatIndex, f.getFormatRecord());

      return f;
    }
    catch (NumFormatRecordsException e)
    {
      logger.warn("Maximum number of format records exceeded.  Using " +
                  "default format.");

      return WritableWorkbook.NORMAL_STYLE;
    }
  }


  /**
   * Accessor for the maximum column outline level
   *
   * @return the maximum column outline level, or 0 if no outlines/groups
   */
  public int getMaxColumnOutlineLevel()
  {
    return maxColumnOutlineLevel;
  }

  /**
   * Accessor for the maximum row outline level
   *
   * @return the maximum row outline level, or 0 if no outlines/groups
   */
  public int getMaxRowOutlineLevel()
  {
    return maxRowOutlineLevel;
  }
}
