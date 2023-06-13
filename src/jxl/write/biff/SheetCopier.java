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
import jxl.Hyperlink;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Range;
import jxl.Sheet;
import jxl.WorkbookSettings;
import jxl.biff.AutoFilter;
import jxl.biff.CellReferenceHelper;
import jxl.biff.ConditionalFormat;
import jxl.biff.DataValidation;
import jxl.biff.FormattingRecords;
import jxl.biff.FormulaData;
import jxl.biff.NumFormatRecordsException;
import jxl.biff.SheetRangeImpl;
import jxl.biff.XFRecord;
import jxl.biff.drawing.Chart;
import jxl.biff.drawing.ComboBox;
import jxl.biff.drawing.DrawingGroupObject;
import jxl.format.CellFormat;
import jxl.biff.formula.FormulaException;
import jxl.read.biff.SheetImpl;
import jxl.read.biff.NameRecord;
import jxl.read.biff.WorkbookParser;
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
class SheetCopier
{
  private static final Logger LOGGER = Logger.getLogger(SheetCopier.class);

  private final SheetImpl fromSheet;
  private final WritableSheetImpl toSheet;
  private final WorkbookSettings workbookSettings;

  // Objects used by the sheet
  private TreeSet<ColumnInfoRecord> columnFormats;
  private FormattingRecords formatRecords;
  private ArrayList<WritableHyperlink> hyperlinks;
  private MergedCells mergedCells;
  private final HorizontalPageBreaksRecord rowBreaks;
  private final VerticalPageBreaksRecord columnBreaks;
  private SheetWriter sheetWriter;
  private ArrayList<DrawingGroupObject> drawings;
  private ArrayList<WritableImage> images;
  private ArrayList<ConditionalFormat> conditionalFormats;
  private ArrayList<WritableCell> validatedCells;
  private AutoFilter autoFilter;
  private DataValidation dataValidation;
  private ComboBox comboBox;
  private PLSRecord plsRecord;
  private boolean chartOnly;
  private ButtonPropertySetRecord buttonPropertySet;
  private int numRows;
  private int maxRowOutlineLevel;
  private int maxColumnOutlineLevel;

  // Objects used to maintain state during the copy process
  private HashMap<Integer, WritableCellFormat> xfRecords;
  private HashMap<Integer, Integer> fonts;
  private HashMap<Integer, Integer> formats;

  public SheetCopier(Sheet f, WritableSheet t,
          HorizontalPageBreaksRecord rb,
          VerticalPageBreaksRecord cb) {
    fromSheet = (SheetImpl) f;
    toSheet = (WritableSheetImpl) t;
    workbookSettings = toSheet.getWorkbook().getSettings();
    chartOnly = false;
    rowBreaks = rb;
    columnBreaks = cb;
  }

  void setColumnFormats(TreeSet<ColumnInfoRecord> cf)
  {
    columnFormats = cf;
  }

  void setFormatRecords(FormattingRecords fr)
  {
    formatRecords = fr;
  }

  void setHyperlinks(ArrayList<WritableHyperlink> h)
  {
    hyperlinks = h;
  }

  void setMergedCells(MergedCells mc)
  {
    mergedCells = mc;
  }

  void setSheetWriter(SheetWriter sw)
  {
    sheetWriter = sw;
  }

  void setDrawings(ArrayList<DrawingGroupObject> d)
  {
    drawings = d;
  }

  void setImages(ArrayList<WritableImage> i)
  {
    images = i;
  }

  void setConditionalFormats(ArrayList<ConditionalFormat> cf)
  {
    conditionalFormats = cf;
  }

  void setValidatedCells(ArrayList<WritableCell> vc)
  {
    validatedCells = vc;
  }

  AutoFilter getAutoFilter()
  {
    return autoFilter;
  }

  DataValidation getDataValidation()
  {
    return dataValidation;
  }

  ComboBox getComboBox()
  {
    return comboBox;
  }

  PLSRecord getPLSRecord()
  {
    return plsRecord;
  }

  boolean isChartOnly()
  {
    return chartOnly;
  }

  ButtonPropertySetRecord getButtonPropertySet()
  {
    return buttonPropertySet;
  }

  /**
   * Copies a sheet from a read-only version to the writable version.
   * Performs shallow copies
   */
  public void copySheet()
  {
    shallowCopyCells();

    // Copy the column info records
    jxl.read.biff.ColumnInfoRecord[] readCirs = fromSheet.getColumnInfos();

    for (jxl.read.biff.ColumnInfoRecord rcir : readCirs)
      for (int j = rcir.getStartColumn(); j <= rcir.getEndColumn() ; j++)
      {
        ColumnInfoRecord cir = new ColumnInfoRecord(rcir, j,
                formatRecords);
        cir.setHidden(rcir.getHidden());
        columnFormats.add(cir);
      }

    // Copy the hyperlinks
    for (Hyperlink hl : fromSheet.getHyperlinks()) {
      WritableHyperlink hr = new WritableHyperlink(hl, toSheet);
      hyperlinks.add(hr);
    }

    // Copy the merged cells
    for (Range range : fromSheet.getMergedCells())
      mergedCells.add(new SheetRangeImpl((SheetRangeImpl) range, toSheet));

    // Copy the row properties
    try
    {
      jxl.read.biff.RowRecord[] rowprops  = fromSheet.getRowProperties();

      for (jxl.read.biff.RowRecord rowprop : rowprops) {
        RowRecord rr = toSheet.getRowRecord(rowprop.getRowNumber());
        XFRecord format = rowprop.hasDefaultFormat() ?
          formatRecords.getXFRecord(rowprop.getXFIndex()) : null;
        rr.setRowDetails(rowprop.getRowHeight(),
                         rowprop.matchesDefaultFontHeight(),
                         rowprop.isCollapsed(),
                         rowprop.getOutlineLevel(),
                         rowprop.getGroupStart(),
                         format);
        numRows = Math.max(numRows, rowprop.getRowNumber() + 1);
      }
    }
    catch (RowsExceededException e)
    {
      // Handle the rows exceeded exception - this cannot occur since
      // the sheet we are copying from will have a valid number of rows
      Assert.verify(false);
    }

    // Copy the headers and footers
    //    sheetWriter.setHeader(new HeaderRecord(si.getHeader()));
    //    sheetWriter.setFooter(new FooterRecord(si.getFooter()));

    // Copy the page breaks
    rowBreaks.setRowBreaks(fromSheet.getRowPageBreaks());
    columnBreaks.setColumnBreaks(fromSheet.getColumnPageBreaks());

    // Copy the charts
    sheetWriter.setCharts(fromSheet.getCharts());

    // Copy the drawings
    DrawingGroupObject[] dr = fromSheet.getDrawings();
    for (DrawingGroupObject dgo : dr)
      if (dgo instanceof jxl.biff.drawing.Drawing) {
        WritableImage wi = new WritableImage
          (dgo, toSheet.getWorkbook().getDrawingGroup());
        drawings.add(wi);
        images.add(wi);
      }
      else if (dgo instanceof jxl.biff.drawing.Comment) {
        jxl.biff.drawing.Comment c = dgo instanceof jxl.biff.drawing.CommentBiff7
                ? new jxl.biff.drawing.CommentBiff7(dgo, toSheet.getWorkbook().getDrawingGroup(), workbookSettings)
                : new jxl.biff.drawing.CommentBiff8(dgo, toSheet.getWorkbook().getDrawingGroup());
        drawings.add(c);
        // Set up the reference on the cell value
        CellValue cv = (CellValue) toSheet.getWritableCell(c.getColumn(),
                c.getRow());
        Assert.verify(cv.getCellFeatures() != null);
        cv.getWritableCellFeatures().setCommentDrawing(c);
      }
      else if (dgo instanceof jxl.biff.drawing.Button) {
        jxl.biff.drawing.Button b =
                new jxl.biff.drawing.Button
                (dgo,
                 toSheet.getWorkbook().getDrawingGroup(),
                 workbookSettings);
        drawings.add(b);
      }
      else if (dgo instanceof jxl.biff.drawing.ComboBox) {
        jxl.biff.drawing.ComboBox cb =
                new jxl.biff.drawing.ComboBox
                (dgo,
                 toSheet.getWorkbook().getDrawingGroup(),
                 workbookSettings);
        drawings.add(cb);
      }
      else if (dgo instanceof jxl.biff.drawing.CheckBox) {
        jxl.biff.drawing.CheckBox cb =
                new jxl.biff.drawing.CheckBox
                (dgo,
                toSheet.getWorkbook().getDrawingGroup(),
                workbookSettings);
        drawings.add(cb);
      }

    // Copy the data validations
    DataValidation rdv = fromSheet.getDataValidation();
    if (rdv != null)
    {
      dataValidation = new DataValidation(rdv,
                                          toSheet.getWorkbook(),
                                          toSheet.getWorkbook(),
                                          workbookSettings);
      int objid = dataValidation.getComboBoxObjectId();

      if (objid != 0)
      {
        comboBox = (ComboBox) drawings.get(objid);
      }
    }

    // Copy the conditional formats
    ConditionalFormat[] cf = fromSheet.getConditionalFormats();
    if (cf.length > 0)
    {
      conditionalFormats.addAll(Arrays.asList(cf));
    }

    // Get the autofilter
    autoFilter = fromSheet.getAutoFilter();

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromSheet.getWorkspaceOptions());

    // Set a flag to indicate if it contains a chart only
    if (fromSheet.getSheetBof().isChart())
    {
      chartOnly = true;
      sheetWriter.setChartOnly();
    }

    // Copy the environment specific print record
    if (fromSheet.getPLS() != null)
    {
      if (fromSheet.getWorkbookBof().isBiff7())
      {
        LOGGER.warn("Cannot copy Biff7 print settings record - ignoring");
      }
      else
      {
        plsRecord = new PLSRecord(fromSheet.getPLS());
      }
    }

    // Copy the button property set
    if (fromSheet.getButtonPropertySet() != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (fromSheet.getButtonPropertySet());
    }

    // Copy the outline levels
    maxRowOutlineLevel = fromSheet.getMaxRowOutlineLevel();
    maxColumnOutlineLevel = fromSheet.getMaxColumnOutlineLevel();
  }

  /**
   * Copies a sheet from a read-only version to the writable version.
   * Performs shallow copies
   */
  public void copyWritableSheet()
  {
    shallowCopyCells();

    /*
    // Copy the column formats
    Iterator cfit = fromWritableSheet.columnFormats.iterator();
    while (cfit.hasNext())
    {
      ColumnInfoRecord cv = new ColumnInfoRecord
        ((ColumnInfoRecord) cfit.next());
      columnFormats.add(cv);
    }

    // Copy the merged cells
    Range[] merged = fromWritableSheet.getMergedCells();

    for (int i = 0; i < merged.length; i++)
    {
      mergedCells.add(new SheetRangeImpl((SheetRangeImpl)merged[i], this));
    }

    // Copy the row properties
    try
    {
      RowRecord[] copyRows = fromWritableSheet.rows;
      RowRecord row = null;
      for (int i = 0; i < copyRows.length ; i++)
      {
        row = copyRows[i];

        if (row != null &&
            (!row.isDefaultHeight() ||
             row.isCollapsed()))
        {
          RowRecord rr = getRowRecord(i);
          rr.setRowDetails(row.getRowHeight(),
                           row.matchesDefaultFontHeight(),
                           row.isCollapsed(),
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
    rowBreaks = new ArrayList(fromWritableSheet.rowBreaks);

    // Copy the vertical page breaks
    columnBreaks = new ArrayList(fromWritableSheet.columnBreaks);

    // Copy the data validations
    DataValidation rdv = fromWritableSheet.dataValidation;
    if (rdv != null)
    {
      dataValidation = new DataValidation(rdv,
                                          workbook,
                                          workbook,
                                          workbookSettings);
    }

    // Copy the charts
    sheetWriter.setCharts(fromWritableSheet.getCharts());

    // Copy the drawings
    DrawingGroupObject[] dr = si.getDrawings();
    for (int i = 0 ; i < dr.length ; i++)
    {
      if (dr[i] instanceof jxl.biff.drawing.Drawing)
      {
        WritableImage wi = new WritableImage(dr[i],
                                             workbook.getDrawingGroup());
        drawings.add(wi);
        images.add(wi);
      }

      // Not necessary to copy the comments, as they will be handled by
      // the deep copy of the individual cells
    }

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromWritableSheet.getWorkspaceOptions());

    // Copy the environment specific print record
    if (fromWritableSheet.plsRecord != null)
    {
      plsRecord = new PLSRecord(fromWritableSheet.plsRecord);
    }

    // Copy the button property set
    if (fromWritableSheet.buttonPropertySet != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (fromWritableSheet.buttonPropertySet);
    }
    */
  }

  /**
   * Imports a sheet from a different workbook, doing a deep copy
   */
  public void importSheet()
  {
    xfRecords = new HashMap<>();
    fonts = new HashMap<>();
    formats = new HashMap<>();

    deepCopyCells();

    // Copy the column info records
    jxl.read.biff.ColumnInfoRecord[] readCirs = fromSheet.getColumnInfos();

    for (jxl.read.biff.ColumnInfoRecord rcir : readCirs)
      for (int j = rcir.getStartColumn(); j <= rcir.getEndColumn() ; j++)
      {
        ColumnInfoRecord cir = new ColumnInfoRecord(rcir, j);
        int xfIndex = cir.getXfIndex();
        XFRecord cf = xfRecords.get(xfIndex);

        if (cf == null)
        {
          CellFormat readFormat = fromSheet.getColumnView(j).getFormat();
          WritableCellFormat wcf = copyCellFormat(readFormat);
        }

        cir.setCellFormat(cf);
        cir.setHidden(rcir.getHidden());
        columnFormats.add(cir);
      }

    // Copy the hyperlinks
    for (Hyperlink hl : fromSheet.getHyperlinks()) {
      WritableHyperlink hr = new WritableHyperlink(hl, toSheet);
      hyperlinks.add(hr);
    }

    // Copy the merged cells
    for (Range range : fromSheet.getMergedCells())
      mergedCells.add(new SheetRangeImpl((SheetRangeImpl) range, toSheet));

    // Copy the row properties
    try
    {
      jxl.read.biff.RowRecord[] rowprops  = fromSheet.getRowProperties();

      for (jxl.read.biff.RowRecord rowprop : rowprops) {
        RowRecord rr = toSheet.getRowRecord(rowprop.getRowNumber());
        XFRecord format = null;
        jxl.read.biff.RowRecord rowrec = rowprop;
        if (rowrec.hasDefaultFormat())
        {
          format = xfRecords.get(rowrec.getXFIndex());

          if (format == null)
          {
            int rownum = rowrec.getRowNumber();
            CellFormat readFormat = fromSheet.getRowView(rownum).getFormat();
            WritableCellFormat wcf = copyCellFormat(readFormat);
          }
        }
        rr.setRowDetails(rowrec.getRowHeight(),
                rowrec.matchesDefaultFontHeight(),
                rowrec.isCollapsed(),
                rowrec.getOutlineLevel(),
                rowrec.getGroupStart(),
                format);
        numRows = Math.max(numRows, rowprop.getRowNumber() + 1);
      }
    }
    catch (RowsExceededException e)
    {
      // Handle the rows exceeded exception - this cannot occur since
      // the sheet we are copying from will have a valid number of rows
      Assert.verify(false);
    }

    // Copy the headers and footers
    //    sheetWriter.setHeader(new HeaderRecord(si.getHeader()));
    //    sheetWriter.setFooter(new FooterRecord(si.getFooter()));

    // Copy the page breaks
    rowBreaks.setRowBreaks(fromSheet.getRowPageBreaks());
    columnBreaks.setColumnBreaks(fromSheet.getColumnPageBreaks());

    // Copy the charts
    Chart[] fromCharts = fromSheet.getCharts();
    if (fromCharts != null && fromCharts.length > 0)
    {
      LOGGER.warn("Importing of charts is not supported");
      /*
      sheetWriter.setCharts(fromSheet.getCharts());
      IndexMapping xfMapping = new IndexMapping(200);
      for (Iterator i = xfRecords.keySet().iterator(); i.hasNext();)
      {
        Integer key = (Integer) i.next();
        XFRecord xfmapping = (XFRecord) xfRecords.get(key);
        xfMapping.setMapping(key.intValue(), xfmapping.getXFIndex());
      }

      IndexMapping fontMapping = new IndexMapping(200);
      for (Iterator i = fonts.keySet().iterator(); i.hasNext();)
      {
        Integer key = (Integer) i.next();
        Integer fontmap = (Integer) fonts.get(key);
        fontMapping.setMapping(key.intValue(), fontmap.intValue());
      }

      IndexMapping formatMapping = new IndexMapping(200);
      for (Iterator i = formats.keySet().iterator(); i.hasNext();)
      {
        Integer key = (Integer) i.next();
        Integer formatmap = (Integer) formats.get(key);
        formatMapping.setMapping(key.intValue(), formatmap.intValue());
      }

      // Now reuse the rationalization feature on each chart  to
      // handle the new fonts
      for (int i = 0; i < fromCharts.length ; i++)
      {
        fromCharts[i].rationalize(xfMapping, fontMapping, formatMapping);
      }
      */
    }

    // Copy the drawings
    DrawingGroupObject[] dr = fromSheet.getDrawings();

    // Make sure the destination workbook has a drawing group
    // created in it
    if (dr.length > 0 &&
        toSheet.getWorkbook().getDrawingGroup() == null)
    {
      toSheet.getWorkbook().createDrawingGroup();
    }

    for (DrawingGroupObject dgo : dr)
      if (dgo instanceof jxl.biff.drawing.Drawing) {
        WritableImage wi = new WritableImage(dgo.getX(), dgo.getY(), dgo.getWidth(), dgo.getHeight(), dgo.getImageData());
        toSheet.getWorkbook().addDrawing(wi);
        drawings.add(wi);
        images.add(wi);
      }
      else if (dgo instanceof jxl.biff.drawing.Comment) {
        jxl.biff.drawing.Comment c = dgo instanceof jxl.biff.drawing.CommentBiff7
                ? new jxl.biff.drawing.CommentBiff7(dgo, toSheet.getWorkbook().getDrawingGroup(), workbookSettings)
                : new jxl.biff.drawing.CommentBiff8(dgo, toSheet.getWorkbook().getDrawingGroup());
        drawings.add(c);
        // Set up the reference on the cell value
        CellValue cv = (CellValue) toSheet.getWritableCell(c.getColumn(),
                c.getRow());
        Assert.verify(cv.getCellFeatures() != null);
        cv.getWritableCellFeatures().setCommentDrawing(c);
      }
      else if (dgo instanceof jxl.biff.drawing.Button) {
        jxl.biff.drawing.Button b = new jxl.biff.drawing.Button(dgo, toSheet.getWorkbook().getDrawingGroup(), workbookSettings);
        drawings.add(b);
      }
      else if (dgo instanceof jxl.biff.drawing.ComboBox) {
        jxl.biff.drawing.ComboBox cb = new jxl.biff.drawing.ComboBox(dgo, toSheet.getWorkbook().getDrawingGroup(), workbookSettings);
        drawings.add(cb);
      }

    // Copy the data validations
    DataValidation rdv = fromSheet.getDataValidation();
    if (rdv != null)
    {
      dataValidation = new DataValidation(rdv,
                                          toSheet.getWorkbook(),
                                          toSheet.getWorkbook(),
                                          workbookSettings);
      int objid = dataValidation.getComboBoxObjectId();
      if (objid != 0)
      {
        comboBox = (ComboBox) drawings.get(objid);
      }
    }

    // Copy the workspace options
    sheetWriter.setWorkspaceOptions(fromSheet.getWorkspaceOptions());

    // Set a flag to indicate if it contains a chart only
    if (fromSheet.getSheetBof().isChart())
    {
      chartOnly = true;
      sheetWriter.setChartOnly();
    }

    // Copy the environment specific print record
    if (fromSheet.getPLS() != null)
    {
      if (fromSheet.getWorkbookBof().isBiff7())
      {
        LOGGER.warn("Cannot copy Biff7 print settings record - ignoring");
      }
      else
      {
        plsRecord = new PLSRecord(fromSheet.getPLS());
      }
    }

    // Copy the button property set
    if (fromSheet.getButtonPropertySet() != null)
    {
      buttonPropertySet = new ButtonPropertySetRecord
        (fromSheet.getButtonPropertySet());
    }

    importNames();

    // Copy the outline levels
    maxRowOutlineLevel = fromSheet.getMaxRowOutlineLevel();
    maxColumnOutlineLevel = fromSheet.getMaxColumnOutlineLevel();
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

    if (c instanceof ReadFormulaRecord)
    {
      ReadFormulaRecord rfr = (ReadFormulaRecord) c;
      boolean crossSheetReference = !rfr.handleImportedCellReferences
        (fromSheet.getWorkbook(),
         fromSheet.getWorkbook(),
         workbookSettings);

      if (crossSheetReference)
      {
        try
        {
        LOGGER.warn("Formula " + rfr.getFormula() +
                    " in cell " +
                    CellReferenceHelper.getCellReference(cell.getColumn(),
                                                         cell.getRow()) +
                    " cannot be imported because it references another " +
                    " sheet from the source workbook");
        }
        catch (FormulaException e)
        {
          LOGGER.warn("Formula  in cell " +
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
    {
      wcf = copyCellFormat(cf);
    }

    c.setCellFormat(wcf);

    return c;
  }

  /**
   * Perform a shallow copy of the cells from the specified sheet into this one
   */
  void shallowCopyCells()
  {
    // Copy the cells
    int cells = fromSheet.getRows();
    for (int i = 0;  i < cells; i++)
    {
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
            if (c.getCellFeatures() != null &&
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
   * Perform a deep copy of the cells from the specified sheet into this one
   */
  void deepCopyCells()
  {
    // Copy the cells
    int rowCount = fromSheet.getRows();
    for (int i = 0;  i < rowCount; i++)
    {
      Cell[] row = fromSheet.getRow(i);

      for (Cell cell : row) {
        WritableCell c = deepCopyCell(cell);
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
            if (c.getCellFeatures() != null &&
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
      LOGGER.warn("Maximum number of format records exceeded.  Using " +
                  "default format.");

      return WritableWorkbook.NORMAL_STYLE;
    }
  }

  /**
   * Imports any names defined on the source sheet to the destination workbook
   */
  private void importNames()
  {
    WorkbookParser fromWorkbook = fromSheet.getWorkbook();
    WritableWorkbook toWorkbook = toSheet.getWorkbook();
    int fromSheetIndex = fromWorkbook.getIndex(fromSheet);
    List<NameRecord> nameRecords = fromWorkbook.getNameRecords();
    var rangeNames = toWorkbook.getRangeNames();

    for (var nameRecord : nameRecords) {
      NameRecord.NameRange[] nameRanges = nameRecord.getRanges();

      for (NameRecord.NameRange nameRange : nameRanges) {
        int nameSheetIndex = fromWorkbook.getExternalSheetIndex(nameRange.getExternalSheet());

        if (fromSheetIndex == nameSheetIndex) {
          String name = nameRecord.getName();
          if (rangeNames.contains(name))
            LOGGER.warn("Named range " + name +
                    " is already present in the destination workbook");
          else
            toWorkbook.addNameArea(
                    name,
                    toSheet,
                    nameRange.getFirstColumn(),
                    nameRange.getFirstRow(),
                    nameRange.getLastColumn(),
                    nameRange.getLastRow());
        }
      }
    }
  }

  /**
   * Gets the number of rows - allows for the case where formatting has
   * been applied to rows, even though the row has no data
   *
   * @return the number of rows
   */
  int getRows()
  {
    return numRows;
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
