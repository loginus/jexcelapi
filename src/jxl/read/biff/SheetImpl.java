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

package jxl.read.biff;

import java.util.*;
import java.util.regex.Pattern;

import jxl.*;
import jxl.biff.*;
import jxl.biff.CellReferenceHelper;
import jxl.biff.drawing.*;
import jxl.format.CellFormat;

/**
 * Represents a sheet within a workbook.  Provides a handle to the individual
 * cells, or lines of cells (grouped by Row or Column)
 * In order to simplify this class due to code bloat, the actual reading
 * logic has been delegated to the SheetReaderClass.  This class' main
 * responsibility is now to implement the API methods declared in the
 * Sheet interface
 */
public class SheetImpl implements Sheet
{
  /**
   * The excel file
   */
  private final File excelFile;
  /**
   * A handle to the shared string table
   */
  private final SSTRecord sharedStrings;

  /**
   * A handle to the sheet BOF record, which indicates the stream type
   */
  private final BOFRecord sheetBof;

  /**
   * A handle to the workbook BOF record, which indicates the stream type
   */
  private final BOFRecord workbookBof;

  /**
   * A handle to the formatting records
   */
  private final FormattingRecords formattingRecords;

  /**
   * The name of this sheet
   */
  private String name;

  /**
   * The  number of rows
   */
  private int numRows;

  /**
   * The number of columns
   */
  private int numCols;

  /**
   * The cells
   */
  private final Map<CellCoordinate, Cell> cells = new HashMap<>();

  /**
   * The start position in the stream of this sheet
   */
  private int startPosition;

  /**
   * The list of specified (ie. non default) column widths
   */
  private ColumnInfoRecord[] columnInfos;

  /**
   * The array of row records
   */
  private RowRecord[] rowRecords;

  /**
   * The list of non-default row properties
   */
  private List<RowRecord> rowProperties;

  /**
   * An array of column info records.  They are held this way before
   * they are transferred to the more convenient array
   */
  private List<ColumnInfoRecord> columnInfosArray;

  /**
   * A list of shared formula groups
   */
  private final List<SharedFormulaRecord> sharedFormulas;

  /**
   * A list of hyperlinks on this page
   */
  private List<Hyperlink> hyperlinks;

  /**
   * A list of charts on this page
   */
  private List<Chart> charts;

  /**
   * A list of drawings on this page
   */
  private List<DrawingGroupObject> drawings;

  /**
   * A list of drawings (as opposed to comments/validation/charts) on this
   * page
   */
  private List<Image> images;

  /**
   * A list of data validations on this page
   */
  private DataValidation dataValidation;

  /**
   * A list of merged cells on this page
   */
  private List<Range> mergedCells;

  /**
   * Indicates whether the columnInfos array has been initialized
   */
  private boolean columnInfosInitialized;

  /**
   * Indicates whether the rowRecords array has been initialized
   */
  private boolean rowRecordsInitialized;

  /**
   * Indicates whether or not the dates are based around the 1904 date system
   */
  private final boolean nineteenFour;

  /**
   * The workspace options
   */
  private WorkspaceInformationRecord workspaceOptions;

  /**
   * The hidden flag
   */
  private boolean hidden;

  /**
   * The environment specific print record
   */
  private PLSRecord plsRecord;

  /**
   * The property set record associated with this workbook
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * The sheet settings
   */
  private SheetSettings settings;

  /**
   * The horizontal page breaks contained on this sheet
   */
  private HorizontalPageBreaksRecord rowBreaks;

  /**
   * The vertical page breaks contained on this sheet
   */
  private VerticalPageBreaksRecord columnBreaks;

  /**
   * The maximum row outline level
   */
  private int maxRowOutlineLevel;

  /**
   * The maximum column outline level
   */
  private int maxColumnOutlineLevel;

  /**
   * The list of local names for this sheet
   */
  private List<NameRecord> localNames;

  /**
   * The list of conditional formats for this sheet
   */
  private List<ConditionalFormat> conditionalFormats;

  /**
   * The autofilter information
   */
  private AutoFilter autoFilter;

  /**
   * A handle to the workbook which contains this sheet.  Some of the records
   * need this in order to reference external sheets
   */
  private final WorkbookParser workbook;

  /**
   * A handle to the workbook settings
   */
  private final WorkbookSettings workbookSettings;

  /**
   * Constructor
   *
   * @param f the excel file
   * @param sst the shared string table
   * @param fr formatting records
   * @param sb the bof record which indicates the start of the sheet
   * @param wb the bof record which indicates the start of the sheet
   * @param nf the 1904 flag
   * @param wp the workbook which this sheet belongs to
   * @exception BiffException
   */
  SheetImpl(File f,
            SSTRecord sst,
            FormattingRecords fr,
            BOFRecord sb,
            BOFRecord wb,
            boolean nf,
            WorkbookParser wp)
    throws BiffException
  {
    excelFile = f;
    sharedStrings = sst;
    formattingRecords = fr;
    sheetBof = sb;
    workbookBof = wb;
    columnInfosArray = new ArrayList<>();
    sharedFormulas = new ArrayList<>();
    hyperlinks = new ArrayList<>();
    rowProperties = new ArrayList<>(10);
    columnInfosInitialized = false;
    rowRecordsInitialized = false;
    nineteenFour = nf;
    workbook = wp;
    workbookSettings = workbook.getSettings();

    // Mark the position in the stream, and then skip on until the end
    startPosition = f.getPos();

    if (sheetBof.isChart())
    {
      // Set the start pos to include the bof so the sheet reader can handle it
      startPosition -= (sheetBof.getLength() + 4);
    }

    int bofs = 1;

    while (bofs >= 1)
    {
      Record r = f.next();

      // use this form for quick performance
      if (r.getCode() == Type.EOF.value)
      {
        bofs--;
      }

      if (r.getCode() == Type.BOF.value)
      {
        bofs++;
      }
    }
  }

  /**
   * Returns the cell for the specified location eg. "A4", using the
   * CellReferenceHelper
   *
   * @param loc the cell reference
   * @return the cell at the specified co-ordinates
   */
  @Override
  public Cell getCell(String loc)
  {
    return getCell(CellReferenceHelper.getColumn(loc),
                   CellReferenceHelper.getRow(loc));
  }

  /**
   * Returns the cell specified at this row and at this column
   *
   * @param row the row number
   * @param column the column number
   * @return the cell at the specified co-ordinates
   */
  @Override
  public Cell getCell(int column, int row)
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells.isEmpty())
      readSheet();

    return cells.computeIfAbsent(new CellCoordinate(column, row), EmptyCell::new);
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  @Override
  public Cell findCell(String contents)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findCell(contents);
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   *
   * @param contents the string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the range
   * @param lastCol the last column within the range
   * @param lastRow the last row within the range
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  @Override
  public Cell findCell(String contents,
                       int firstCol,
                       int firstRow,
                       int lastCol,
                       int lastRow,
                       boolean reverse)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findCell(contents,
                               firstCol,
                               firstRow,
                               lastCol,
                               lastRow,
                               reverse);
  }

  /**
   * Gets the cell whose contents match the regular expressionstring passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform
   *
   * @param pattern the regular expression string to match
   * @param firstCol the first column within the range
   * @param firstRow the first row of the range
   * @param lastRow the last row within the range
   * @param lastCol the last column within the ranage
   * @param reverse indicates whether to perform a reverse search or not
   * @return the Cell whose contents match the parameter, null if not found
   */
  @Override
  public Cell findCell(Pattern pattern,
                       int firstCol,
                       int firstRow,
                       int lastCol,
                       int lastRow,
                       boolean reverse)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findCell(pattern,
                               firstCol,
                               firstRow,
                               lastCol,
                               lastRow,
                               reverse);
  }

  /**
   * Gets the cell whose contents match the string passed in.
   * If no match is found, then null is returned.  The search is performed
   * on a row by row basis, so the lower the row number, the more
   * efficiently the algorithm will perform.  This method differs
   * from the findCell methods in that only cells with labels are
   * queried - all numerical cells are ignored.  This should therefore
   * improve performance.
   *
   * @param  contents the string to match
   * @return the Cell whose contents match the paramter, null if not found
   */
  @Override
  public LabelCell findLabelCell(String contents)
  {
    CellFinder cellFinder = new CellFinder(this);
    return cellFinder.findLabelCell(contents);
  }

  /**
   * Returns the number of rows in this sheet
   *
   * @return the number of rows in this sheet
   */
  @Override
  public int getRows()
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells.isEmpty())
      readSheet();

    return numRows;
  }

  /**
   * Returns the number of columns in this sheet
   *
   * @return the number of columns in this sheet
   */
  @Override
  public int getColumns()
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells.isEmpty())
      readSheet();

    return numCols;
  }

  /**
   * Gets all the cells on the specified row.  The returned array will
   * be stripped of all trailing empty cells
   *
   * @param row the rows whose cells are to be returned
   * @return the cells on the given row
   */
  @Override
  public Cell[] getRow(int row)
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells.isEmpty())
      readSheet();

    int col = cells.keySet().stream()
            .filter(coord -> coord.getRow() == row)
            .mapToInt(CellCoordinate::getColumn)
            .max()
            .orElse(-1);

    // Only create entries for non-null cells
    Cell[] c = new Cell[col + 1];

    for (int i = 0; i <= col; i++)
    {
      c[i] = getCell(i, row);
    }
    return c;
  }

  /**
   * Gets all the cells on the specified column.  The returned array
   * will be stripped of all trailing empty cells
   *
   * @param col the column whose cells are to be returned
   * @return the cells on the specified column
   */
  @Override
  public Cell[] getColumn(int col)
  {
    // just in case this has been cleared, but something else holds
    // a reference to it
    if (cells.isEmpty())
      readSheet();

    int row = cells.keySet().stream()
            .filter(coord -> coord.getColumn() == col)
            .mapToInt(CellCoordinate::getRow)
            .max()
            .orElse(-1);

    // Only create entries for non-null cells
    Cell[] c = new Cell[row + 1];

    for (int i = 0; i <= row; i++)
    {
      c[i] = getCell(col, i);
    }
    return c;
  }

  /**
   * Gets the name of this sheet
   *
   * @return the name of the sheet
   */
  @Override
  public String getName()
  {
    return name;
  }

  /**
   * Sets the name of this sheet
   *
   * @param s the sheet name
   */
  final void setName(String s)
  {
    name = s;
  }

  /**
   * Gets the column info record for the specified column.  If no
   * column is specified, null is returned
   *
   * @param col the column
   * @return the ColumnInfoRecord if specified, NULL otherwise
   */
  public ColumnInfoRecord getColumnInfo(int col)
  {
    if (!columnInfosInitialized)
    {
      // Initialize the array
      for (ColumnInfoRecord cir : columnInfosArray) {
        int startcol = Math.max(0, cir.getStartColumn());
        int endcol = Math.min(columnInfos.length - 1, cir.getEndColumn());

        for (int c = startcol; c <= endcol; c++)
        {
          columnInfos[c] = cir;
        }

        if (endcol < startcol)
        {
          columnInfos[startcol] = cir;
        }
      }

      columnInfosInitialized = true;
    }

    return col < columnInfos.length ? columnInfos[col] : null;
  }

  /**
   * Gets all the column info records
   *
   * @return the ColumnInfoRecordArray
   */
  public ColumnInfoRecord[] getColumnInfos()
  {
    // Just chuck all the column infos we have into an array
    return columnInfosArray.toArray(ColumnInfoRecord[]::new);
  }

  /**
   * Sets the visibility of this sheet
   *
   * @param h hidden flag
   */
  final void setHidden(boolean h)
  {
    hidden = h;
  }

  /**
   * Clears out the array of cells.  This is done for memory allocation
   * reasons when reading very large sheets
   */
  final void clear()
  {
    cells.clear();
    mergedCells = null;
    columnInfosArray.clear();
    sharedFormulas.clear();
    hyperlinks.clear();
    columnInfosInitialized = false;

    if (!workbookSettings.getGCDisabled())
    {
      System.gc();
    }
  }

  /**
   * Reads in the contents of this sheet
   */
  final void readSheet()
  {
    SheetReader reader = new SheetReader(excelFile,
                                         sharedStrings,
                                         formattingRecords,
                                         sheetBof,
                                         workbookBof,
                                         nineteenFour,
                                         workbook,
                                         startPosition,
                                         this);
    reader.read();

    // Take stuff that was read in
    numRows = reader.getNumRows();
    numCols = reader.getNumCols();
    cells.putAll(reader.getCells());
    rowProperties = reader.getRowProperties();
    columnInfosArray = reader.getColumnInfosArray();
    hyperlinks = reader.getHyperlinks();
    conditionalFormats = reader.getConditionalFormats();
    autoFilter = reader.getAutoFilter();
    charts = reader.getCharts();
    drawings = reader.getDrawings();
    dataValidation = reader.getDataValidation();
    mergedCells = reader.getMergedCells();
    settings = reader.getSettings();
    settings.setHidden(hidden);
    rowBreaks = reader.getRowBreaks();
    columnBreaks = reader.getColumnBreaks();
    workspaceOptions = reader.getWorkspaceOptions();
    plsRecord = reader.getPLS();
    buttonPropertySet = reader.getButtonPropertySet();
    maxRowOutlineLevel = reader.getMaxRowOutlineLevel();
    maxColumnOutlineLevel = reader.getMaxColumnOutlineLevel();

    if (!workbookSettings.getGCDisabled())
    {
      System.gc();
    }

    if (!columnInfosArray.isEmpty())
    {
      ColumnInfoRecord cir = columnInfosArray.get(columnInfosArray.size() - 1);
      columnInfos = new ColumnInfoRecord[cir.getEndColumn() + 1];
    }
    else
    {
      columnInfos = new ColumnInfoRecord[0];
    }

    // Add any local names
    if (localNames != null)
    {
      for (NameRecord nr : localNames) {
        if (nr.getBuiltInName() == BuiltInName.PRINT_AREA)
        {
          if(nr.getRanges().length > 0)
          {
            NameRecord.NameRange rng = nr.getRanges()[0];
            settings.setPrintArea(rng.getFirstColumn(),
                                  rng.getFirstRow(),
                                  rng.getLastColumn(),
                                  rng.getLastRow());
          }
        }
        else if (nr.getBuiltInName() == BuiltInName.PRINT_TITLES)
       	{
          // There can be 1 or 2 entries.
          // Row entries have hardwired column entries (first and last
          //  possible column)
          // Column entries have hardwired row entries (first and last
          // possible row)
          for (NameRecord.NameRange rng : nr.getRanges())
            if (rng.getFirstColumn() == 0 && rng.getLastColumn() == 255)
            {
              settings.setPrintTitlesRow(rng.getFirstRow(),
                      rng.getLastRow());
            }
            else
            {
              settings.setPrintTitlesCol(rng.getFirstColumn(),
                      rng.getLastColumn());
            }
        }
      }
    }
  }

  /**
   * Gets the hyperlinks on this sheet
   *
   * @return an array of hyperlinks
   */
  @Override
  public Hyperlink[] getHyperlinks()
  {
    return hyperlinks.toArray(Hyperlink[]::new);
  }

  /**
   * Gets the cells which have been merged on this sheet
   *
   * @return a List of range objects
   */
  @Override
  public List<Range> getMergedCells()
  {
    if (mergedCells == null)
      return List.of();

    return mergedCells;
  }

  /**
   * Gets the non-default rows.  Used when copying spreadsheets
   *
   * @return an array of row properties
   */
  public RowRecord[] getRowProperties() {
    return rowProperties.toArray(RowRecord[]::new);
  }

  /**
   * Gets the data validations.  Used when copying sheets
   *
   * @return the data validations
   */
  public DataValidation getDataValidation()
  {
    return dataValidation;
  }

  /**
   * Gets the row record.  Usually called by the cell in the specified
   * row in order to determine its size
   *
   * @param r the row
   * @return the RowRecord for the specified row
   */
  RowRecord getRowInfo(int r)
  {
    if (!rowRecordsInitialized)
    {
      rowRecords = new RowRecord[getRows()];

      for (RowRecord rr : rowProperties) {
        int rownum = rr.getRowNumber();
        if (rownum < rowRecords.length)
          rowRecords[rownum] = rr;
      }

      rowRecordsInitialized = true;
    }

    return r < rowRecords.length ? rowRecords[r] : null;
  }

  /**
   * Gets the row breaks.  Called when copying sheets
   *
   * @return the explicit row breaks
   */
  @Override
  public final IHorizontalPageBreaks getRowPageBreaks()
  {
    return rowBreaks;
  }

  /**
   * Gets the row breaks.  Called when copying sheets
   *
   * @return the explicit row breaks
   */
  @Override
  public final IVerticalPageBreaks getColumnPageBreaks()
  {
    return columnBreaks;
  }

  /**
   * Gets the charts.  Called when copying sheets
   *
   * @return the charts on this page
   */
  public final Chart[] getCharts()
  {
    return charts.toArray(Chart[]::new);
  }

  /**
   * Gets the drawings.  Called when copying sheets
   *
   * @return the drawings on this page
   */
  public final DrawingGroupObject[] getDrawings()
  {
    return drawings.toArray(DrawingGroupObject[]::new);
  }

  /**
   * Gets the workspace options for this sheet.  Called during the copy
   * process
   *
   * @return the workspace options
   */
  public WorkspaceInformationRecord getWorkspaceOptions()
  {
    return workspaceOptions;
  }

  /**
   * Accessor for the sheet settings
   *
   * @return the settings for this sheet
   */
  @Override
  public SheetSettings getSettings()
  {
    return settings;
  }

  /**
   * Accessor for the workbook.  In addition to be being used by this package,
   * it is also used during the importSheet process
   *
   * @return  the workbook
   */
  public WorkbookParser getWorkbook()
  {
    return workbook;
  }

  /**
   * Gets the column format for the specified column
   *
   * @param col the column number
   * @return the column format, or NULL if the column has no specific format
   * @deprecated use getColumnView instead
   */
  @Override
  public CellFormat getColumnFormat(int col)
  {
    CellView cv = getColumnView(col);
    return cv.getFormat();
  }

  /**
   * Gets the column width for the specified column
   *
   * @param col the column number
   * @return the column width, or the default width if the column has no
   *         specified format
   */
  @Override
  public int getColumnWidth(int col)
  {
    return getColumnView(col).getSize() / 256;
  }

  /**
   * Gets the column width for the specified column
   *
   * @param col the column number
   * @return the column format, or the default format if no override is
             specified
   */
  @Override
  public CellView getColumnView(int col)
  {
    ColumnInfoRecord cir = getColumnInfo(col);
    CellView cv = new CellView();

    if (cir != null)
    {
      cv.setDimension(cir.getWidth() / 256); //deprecated
      cv.setSize(cir.getWidth());
      cv.setHidden(cir.getHidden());
      cv.setFormat(formattingRecords.getXFRecord(cir.getXFIndex()));
    }
    else
    {
      cv.setDimension(settings.getDefaultColumnWidth()); //deprecated
      cv.setSize(settings.getDefaultColumnWidth() * 256);
    }

    return cv;
  }

  /**
   * Gets the row height for the specified column
   *
   * @param row the row number
   * @return the row height, or the default height if the row has no
   *         specified format
   * @deprecated use getRowView instead
   */
  @Override
  public int getRowHeight(int row)
  {
    return getRowView(row).getDimension();
  }

  /**
   * Gets the row view for the specified row
   *
   * @param row the row number
   * @return the row format, or the default format if no override is
             specified
   */
  @Override
  public CellView getRowView(int row)
  {
    RowRecord rr = getRowInfo(row);

    CellView cv = new CellView();

    if (rr != null)
    {
      cv.setDimension(rr.getRowHeight()); //deprecated
      cv.setSize(rr.getRowHeight());
      cv.setHidden(rr.isCollapsed());
      if (rr.hasDefaultFormat())
      {
        cv.setFormat(formattingRecords.getXFRecord(rr.getXFIndex()));
      }
    }
    else
    {
      cv.setDimension(settings.getDefaultRowHeight());
      cv.setSize(settings.getDefaultRowHeight()); //deprecated
    }

    return cv;
  }


  /**
   * Used when copying sheets in order to determine the type of this sheet
   *
   * @return the BOF Record
   */
  public BOFRecord getSheetBof()
  {
    return sheetBof;
  }

  /**
   * Used when copying sheets in order to determine the type of the containing
   * workboook
   *
   * @return the workbook BOF Record
   */
  public BOFRecord getWorkbookBof()
  {
    return workbookBof;
  }

  /**
   * Accessor for the environment specific print record, invoked when
   * copying sheets
   *
   * @return the environment specific print record
   */
  public PLSRecord getPLS()
  {
    return plsRecord;
  }

  /**
   * Accessor for the button property set, used during copying
   *
   * @return the button property set
   */
  public ButtonPropertySetRecord getButtonPropertySet()
  {
    return buttonPropertySet;
  }

  /**
   * Accessor for the number of images on the sheet
   *
   * @return the number of images on this sheet
   */
  @Override
  public int getNumberOfImages()
  {
    if (images == null)
    {
      initializeImages();
    }

    return images.size();
  }

  /**
   * Accessor for the image
   *
   * @param i the 0 based image number
   * @return  the image at the specified position
   */
  @Override
  public Image getDrawing(int i)
  {
    if (images == null)
    {
      initializeImages();
    }

    return images.get(i);
  }

  /**
   * Initializes the images
   */
  private void initializeImages()
  {
    if (images != null)
    {
      return;
    }

    images = new ArrayList<>();
    DrawingGroupObject[] dgos = getDrawings();

    for (DrawingGroupObject dgo : dgos)
      if (dgo instanceof Drawing)
        images.add((Image) dgo);
  }

  /**
   * Used by one of the demo programs for debugging purposes only
   */
  public DrawingData getDrawingData()
  {
    SheetReader reader = new SheetReader(excelFile,
                                         sharedStrings,
                                         formattingRecords,
                                         sheetBof,
                                         workbookBof,
                                         nineteenFour,
                                         workbook,
                                         startPosition,
                                         this);
    reader.read();
    return reader.getDrawingData();
  }

  /**
   * Adds a local name to this shate
   *
   * @param nr the local name to add
   */
  void addLocalName(NameRecord nr)
  {
    if (localNames == null)
    {
      localNames = new ArrayList<>();
    }

    localNames.add(nr);
  }

  /**
   * Gets the conditional formats
   *
   * @return the conditional formats
   */
  public ConditionalFormat[] getConditionalFormats()
  {
    ConditionalFormat[] formats =
      new ConditionalFormat[conditionalFormats.size()];
    formats = conditionalFormats.toArray(formats);
    return formats;
  }

  /**
   * Returns the autofilter
   *
   * @return the autofilter
   */
  public AutoFilter getAutoFilter()
  {
    return autoFilter;
  }

  /**
   * Accessor for the maximum column outline level.  Used during a copy
   *
   * @return the maximum column outline level, or 0 if no outlines/groups
   */
  public int getMaxColumnOutlineLevel()
  {
    return maxColumnOutlineLevel;
  }

  /**
   * Accessor for the maximum row outline level.  Used during a copy
   *
   * @return the maximum row outline level, or 0 if no outlines/groups
   */
  public int getMaxRowOutlineLevel()
  {
    return maxRowOutlineLevel;
  }

}
