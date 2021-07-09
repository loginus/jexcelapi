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

import java.nio.file.*;
import java.util.*;
import static java.util.stream.Collectors.toList;

import jxl.common.Assert;
import jxl.common.Logger;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.biff.BuiltInName;
import jxl.biff.CellReferenceHelper;
import jxl.biff.EmptyCell;
import jxl.biff.FontRecord;
import jxl.biff.Fonts;
import jxl.biff.FormatRecord;
import jxl.biff.FormattingRecords;
import jxl.biff.NameRangeException;
import jxl.biff.NumFormatRecordsException;
import jxl.biff.PaletteRecord;
import jxl.biff.RangeImpl;
import jxl.biff.StringHelper;
import jxl.biff.Type;
import jxl.biff.WorkbookMethods;
import jxl.biff.XCTRecord;
import jxl.biff.XFRecord;
import jxl.biff.drawing.DrawingGroup;
import jxl.biff.drawing.MsoDrawingGroupRecord;
import jxl.biff.drawing.Origin;
import jxl.biff.formula.ExternalSheet;

/**
 * Parses the biff file passed in, and builds up an internal representation of
 * the spreadsheet
 */
public class WorkbookParser extends Workbook
  implements ExternalSheet, WorkbookMethods
{
  /**
   * The logger
   */
  private static final Logger LOGGER = Logger.getLogger(WorkbookParser.class);

  /**
   * The excel file
   */
  private final File excelFile;

  /**
   * The number of open bofs
   */
  private int bofs;

  /**
   * Indicates whether or not the dates are based around the 1904 date system
   */
  private boolean nineteenFour;

  /**
   * The shared string table
   */
  private SSTRecord sharedStrings;

  /**
   * The names of all the worksheets
   */
  private final List<BoundsheetRecord> boundsheets = new ArrayList<>(10);

  /**
   * The fonts used by this workbook
   */
  private final Fonts fonts = new Fonts();

  /**
   * The xf records
   */
  private final FormattingRecords formattingRecords = new FormattingRecords(fonts);

  /**
   * The sheets contained in this workbook
   */
  private final List<SheetImpl> sheets = new ArrayList<>(10);

  /**
   * The last sheet accessed
   */
  private SheetImpl lastSheet;

  /**
   * The index of the last sheet retrieved
   */
  private int lastSheetIndex = -1;

  /**
   * The named records found in this workbook
   */
  private final Map<String, NameRecord> namedRecords = new HashMap<>();

  /**
   * The list of named records
   */
  private List<NameRecord> nameTable;

  /**
   * The list of add in functions
   */
  private List<String> addInFunctions;

  /**
   * The external sheet record.  Used by formulas, and names
   */
  private ExternalSheetRecord externSheet;

  /**
   * The list of supporting workbooks - used by formulas
   */
  private final List<SupbookRecord> supbooks = new ArrayList<>(10);

  /**
   * The bof record for this workbook
   */
  private BOFRecord workbookBof;

  /**
   * The Mso Drawing Group record for this workbook
   */
  private MsoDrawingGroupRecord msoDrawingGroup;

  /**
   * The property set record associated with this workbook
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * Workbook protected flag
   */
  private boolean wbProtected = false;

  /**
   * Contains macros flag
   */
  private boolean containsMacros = false;

  /**
   * The workbook settings
   */
  private final WorkbookSettings settings;

  /**
   * The drawings contained in this workbook
   */
  private DrawingGroup drawingGroup;

  /**
   * The country record (containing the language and regional settings)
   * for this workbook
   */
  private CountryRecord countryRecord;

  private final List<XCTRecord> xctRecords = new ArrayList<>(10);

  /**
   * Constructs this object from the raw excel data
   *
   * @param f the excel 97 biff file
   * @param s the workbook settings
   */
  public WorkbookParser(File f, WorkbookSettings s)
  {
    super();
    excelFile = f;
    settings = s;
  }

 /**
   * Gets the sheets within this workbook.
   * NOTE:  Use of this method for
   * very large worksheets can cause performance and out of memory problems.
   * Use the alternative method getSheet() to retrieve each sheet individually
   *
   * @return a list of the individual sheets
   */
  @Override
  public List<Sheet> getSheets()
  {
    return Collections.unmodifiableList(sheets);
  }

  /**
   * Interface method from WorkbookMethods - gets the specified
   * sheet within this workbook
   *
   * @param index the zero based index of the required sheet
   * @return The sheet specified by the index
   */
  @Override
  public Sheet getReadSheet(int index)
  {
    return getSheet(index);
  }

  /**
   * Gets the specified sheet within this workbook
   *
   * @param index the zero based index of the required sheet
   * @return The sheet specified by the index
   */
  @Override
  public Sheet getSheet(int index)
  {
    // First see if the last sheet index is the same as this sheet index.
    // If so, then the same sheet is being re-requested, so simply
    // return it instead of rereading it
    if ((lastSheet != null) && lastSheetIndex == index)
    {
      return lastSheet;
    }

    // Flush out all of the cached data in the last sheet
    if (lastSheet != null)
    {
      lastSheet.clear();

      if (!settings.getGCDisabled())
      {
        System.gc();
      }
    }

    lastSheet = sheets.get(index);
    lastSheetIndex = index;
    lastSheet.readSheet();

    return lastSheet;
  }

  /**
   * Gets the sheet with the specified name from within this workbook
   *
   * @param name the sheet name
   * @return The sheet with the specified name, or null if it is not found
   */
  @Override
  public Sheet getSheet(String name)
  {
    // Iterate through the boundsheet records
    int pos = 0;
    boolean found = false;
    Iterator<BoundsheetRecord> i = boundsheets.iterator();

    while (i.hasNext() && !found)
    {
      BoundsheetRecord br = i.next();

      if (br.getName().equals(name))
      {
        found = true;
      }
      else
      {
        pos++;
      }
    }

    return found ? getSheet(pos) : null;
  }

  /**
   * Gets the sheet names
   *
   * @return a list of strings containing the sheet names
   */
  @Override
  public List<String> getSheetNames()
  {
    return boundsheets.stream()
            .map(BoundsheetRecord::getName)
            .collect(toList());
  }


  /**
   * Package protected function which gets the real internal sheet index
   * based upon  the external sheet reference.  This is used for extern sheet
   * references  which are specified in formulas
   *
   * @param index the external sheet reference
   * @return the actual sheet index
   */
  @Override
  public int getExternalSheetIndex(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      return index;
    }

    Assert.verify(externSheet != null);

    int firstTab = externSheet.getFirstTabIndex(index);

    return firstTab;
  }

  /**
   * Package protected function which gets the real internal sheet index
   * based upon  the external sheet reference.  This is used for extern sheet
   * references  which are specified in formulas
   *
   * @param index the external sheet reference
   * @return the actual sheet index
   */
  @Override
  public int getLastExternalSheetIndex(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      return index;
    }

    Assert.verify(externSheet != null);

    int lastTab = externSheet.getLastTabIndex(index);

    return lastTab;
  }

  /**
   * Gets the name of the external sheet specified by the index
   *
   * @param index the external sheet index
   * @return the name of the external sheet
   */
  @Override
  public String getExternalSheetName(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      BoundsheetRecord br = boundsheets.get(index);

      return br.getName();
    }

    int supbookIndex = externSheet.getSupbookIndex(index);
    SupbookRecord sr = supbooks.get(supbookIndex);

    int firstTab = externSheet.getFirstTabIndex(index);
    int lastTab  = externSheet.getLastTabIndex(index);
    String firstTabName;
    String lastTabName;

    if (sr.getType() == SupbookRecord.INTERNAL)
    {
      // It's an internal reference - get the name from the boundsheets list
      if (firstTab == 65535)
      {
         firstTabName = "#REF";
      }
      else
      {
        BoundsheetRecord br = boundsheets.get(firstTab);
        firstTabName = br.getName();
      }

      if (lastTab==65535)
      {
        lastTabName = "#REF";
      }
      else
      {
        BoundsheetRecord br = boundsheets.get(lastTab);
        lastTabName = br.getName();
      }

      String sheetName = (firstTab == lastTab) ? firstTabName :
        firstTabName + ':' + lastTabName;

      // if the sheet name contains apostrophes then escape them
      sheetName = sheetName.indexOf('\'') == -1 ? sheetName :
        StringHelper.replace(sheetName, "\'", "\'\'");


      // if the sheet name contains spaces, then enclose in quotes
      return sheetName.indexOf(' ') == -1 ? sheetName :
        '\'' + sheetName + '\'';
    }
    else if (sr.getType() == SupbookRecord.EXTERNAL)
    {
      // External reference - get the sheet name from the supbook record
      StringBuilder sb = new StringBuilder();
      Path fl = Paths.get(sr.getFileName());
      sb.append("'");
      sb.append(fl.toAbsolutePath().getParent().toString());
      sb.append("[");
      sb.append(fl.getFileName());
      sb.append("]");
      sb.append((firstTab == 65535) ? "#REF" : sr.getSheetName(firstTab));
      if (lastTab != firstTab)
      {
        sb.append(sr.getSheetName(lastTab));
      }
      sb.append("'");
      return sb.toString();
    }

    // An unknown supbook - return unkown
    LOGGER.warn("Unknown Supbook 3");
    return "[UNKNOWN]";
  }

  /**
   * Gets the name of the external sheet specified by the index
   *
   * @param index the external sheet index
   * @return the name of the external sheet
   */
  public String getLastExternalSheetName(int index)
  {
    // For biff7, the whole external reference thing works differently
    // Hopefully for our purposes sheet references will all be local
    if (workbookBof.isBiff7())
    {
      BoundsheetRecord br = boundsheets.get(index);

      return br.getName();
    }

    int supbookIndex = externSheet.getSupbookIndex(index);
    SupbookRecord sr = supbooks.get(supbookIndex);

    int lastTab = externSheet.getLastTabIndex(index);

    if (sr.getType() == SupbookRecord.INTERNAL)
    {
      // It's an internal reference - get the name from the boundsheets list
       if (lastTab == 65535)
       {
         return "#REF";
       }
       else
       {
         BoundsheetRecord br = boundsheets.get(lastTab);
         return br.getName();
       }
    }
    else if (sr.getType() == SupbookRecord.EXTERNAL)
    {
      // External reference - get the sheet name from the supbook record
      StringBuilder sb = new StringBuilder();
      Path fl = Paths.get(sr.getFileName());
      sb.append("'");
      sb.append(fl.toAbsolutePath().getParent().toString());
      sb.append("[");
      sb.append(fl.getFileName());
      sb.append("]");
      sb.append((lastTab == 65535) ? "#REF" : sr.getSheetName(lastTab));
      sb.append("'");
      return sb.toString();
    }

    // An unknown supbook - return unkown
    LOGGER.warn("Unknown Supbook 4");
    return "[UNKNOWN]";
  }

  /**
   * Returns the number of sheets in this workbook
   *
   * @return the number of sheets in this workbook
   */
  @Override
  public int getNumberOfSheets()
  {
    return sheets.size();
  }

  /**
   * Closes this workbook, and frees makes any memory allocated available
   * for garbage collection
   */
  @Override
  public void close()
  {
    if (lastSheet != null)
    {
      lastSheet.clear();
    }
    excelFile.clear();

    if (!settings.getGCDisabled())
    {
      System.gc();
    }
  }

  /**
   * Adds the sheet to the end of the array
   *
   * @param s the sheet to add
   */
  private void addSheet(SheetImpl s)
  {
    sheets.add(s);
  }

  /**
   * Does the hard work of building up the object graph from the excel bytes
   *
   * @exception BiffException
   * @exception PasswordException if the workbook is password protected
   */
  @Override
  protected void parse() throws BiffException, PasswordException
  {
    Record r = null;

    BOFRecord bof = new BOFRecord(excelFile.next());
    workbookBof = bof;
    bofs++;

    if (!bof.isBiff8() && !bof.isBiff7())
    {
      throw new BiffException(BiffException.unrecognizedBiffVersion);
    }

    if (!bof.isWorkbookGlobals())
    {
      throw new BiffException(BiffException.expectedGlobals);
    }
    List<Record> continueRecords = new ArrayList<>();
    List<NameRecord> localNames = new ArrayList<>();
    nameTable = new ArrayList<>();
    addInFunctions = new ArrayList<>();

    // Skip to the first worksheet
    while (bofs == 1)
    {
      r = excelFile.next();

      switch (r.getType()) {
        case SST: {
          // BIFF8 only
          continueRecords.clear();
          Record nextrec = excelFile.peek();
          while (nextrec.getType() == Type.CONTINUE) {
            continueRecords.add(excelFile.next());
            nextrec = excelFile.peek();
          }
          Record[] records = continueRecords.toArray(new Record[continueRecords
                  .size()]);
          sharedStrings = new SSTRecord(r, records);
          break;
        }

        case FILEPASS:
          throw new PasswordException();

        case NAME: {
          NameRecord nr = bof.isBiff8()
                  ? new NameRecord(r, settings, nameTable.size())
                  : new NameRecord(r, settings, nameTable.size(), NameRecord.biff7);
          // Add all local and global names to the name table in order to
          // preserve the indexing
          nameTable.add(nr);
          if (nr.isGlobal())
            namedRecords.put(nr.getName(), nr);
          else
            localNames.add(nr);
          break;
        }

        case FONT: {
          FontRecord fr = bof.isBiff8()
                  ? new FontRecord(r, settings)
                  : new FontRecord(r, settings, FontRecord.biff7);
          fonts.addFont(fr);
          break;
        }

        case PALETTE:
          PaletteRecord palette = new PaletteRecord(r);
          formattingRecords.setPalette(palette);
          break;

        case NINETEENFOUR: {
          NineteenFourRecord nr = new NineteenFourRecord(r);
          nineteenFour = nr.is1904();
          break;
        }

        case FORMAT: {
          FormatRecord fr = bof.isBiff8()
                  ? new FormatRecord(r, settings, FormatRecord.biff8)
                  : new FormatRecord(r, settings, FormatRecord.biff7);
          try {
            formattingRecords.addFormat(fr);
          } catch (NumFormatRecordsException e) {
            // This should not happen.  Bomb out
            Assert.verify(false, e.getMessage());
          }
          break;
        }

        case XF:
          XFRecord xfr = bof.isBiff8()
                  ? new XFRecord(r, settings, XFRecord.biff8)
                  : new XFRecord(r, settings, XFRecord.biff7);
          try {
            formattingRecords.addStyle(xfr);
          } catch (NumFormatRecordsException e) {
            // This should not happen.  Bomb out
            Assert.verify(false, e.getMessage());
          }
          break;

        case BOUNDSHEET:
          BoundsheetRecord br = bof.isBiff8()
                  ? new BoundsheetRecord(r, settings)
                  : new BoundsheetRecord(r, BoundsheetRecord.biff7);

          if (br.isSheet())
            boundsheets.add(br);
          else if (br.isChart() && !settings.getDrawingsDisabled())
            boundsheets.add(br);
          break;

        case EXTERNSHEET:
          if (bof.isBiff8())
            externSheet = new ExternalSheetRecord(r, settings);
          else
            externSheet = new ExternalSheetRecord(r, settings,
                    ExternalSheetRecord.biff7);
          break;

        case XCT:
          xctRecords.add(new XCTRecord(r));
          break;

        case CODEPAGE:
          CodepageRecord cr = new CodepageRecord(r);
          settings.setCharacterSet(cr.getCharacterSet());
          break;

        case SUPBOOK: {
          Record nextrec = excelFile.peek();
          while (nextrec.getType() == Type.CONTINUE) {
            r.addContinueRecord(excelFile.next());
            nextrec = excelFile.peek();
          }
          supbooks.add(new SupbookRecord(r, settings));
          break;

        }
        case EXTERNNAME:
          ExternalNameRecord enr = new ExternalNameRecord(r, settings);
          if (enr.isAddInFunction())
            addInFunctions.add(enr.getName());
          break;

        case PROTECT:
          ProtectRecord pr = new ProtectRecord(r);
          wbProtected = pr.isProtected();
          break;

        case OBJPROJ:
          containsMacros = true;
          break;

        case COUNTRY:
          countryRecord = new CountryRecord(r);
          break;

        case MSODRAWINGGROUP:
          if (!settings.getDrawingsDisabled()) {
            msoDrawingGroup = new MsoDrawingGroupRecord(r);

            if (drawingGroup == null)
              drawingGroup = new DrawingGroup(Origin.READ);

            drawingGroup.add(msoDrawingGroup);

            Record nextrec = excelFile.peek();
            while (nextrec.getType() == Type.CONTINUE) {
              drawingGroup.add(excelFile.next());
              nextrec = excelFile.peek();
            }
          }
          break;

        case BUTTONPROPERTYSET:
          buttonPropertySet = new ButtonPropertySetRecord(r);
          break;

        case EOF:
          bofs--;
          break;

        case REFRESHALL: {
          RefreshAllRecord rfm = new RefreshAllRecord(r);
          settings.setRefreshAll(rfm.getRefreshAll());
          break;
        }

        case TEMPLATE: {
          TemplateRecord rfm = new TemplateRecord(r);
          settings.setTemplate(rfm.getTemplate());
          break;
        }

        case EXCEL9FILE:
          Excel9FileRecord e9f = new Excel9FileRecord(r);
          settings.setExcel9File(e9f.getExcel9File());
          break;

        case WINDOWPROTECT:
          WindowProtectedRecord winp = new WindowProtectedRecord(r);
          settings.setWindowProtected(winp.getWindowProtected());
          break;

        case HIDEOBJ:
          HideobjRecord hobj = new HideobjRecord(r);
          settings.setHideobj(hobj.getHideMode());
          break;

        case WRITEACCESS:
          WriteAccessRecord war = new WriteAccessRecord(r, bof.isBiff8(),
                  settings);
          settings.setWriteAccess(war.getWriteAccess());
          break;
      }
    }

    bof = null;
    if (excelFile.hasNext())
    {
      r = excelFile.next();

      if (r.getType() == Type.BOF)
      {
        bof = new BOFRecord(r);
      }
    }

    // Only get sheets for which there is a corresponding Boundsheet record
    while (bof != null && getNumberOfSheets() < boundsheets.size())
    {
      if (!bof.isBiff8() && !bof.isBiff7())
      {
        throw new BiffException(BiffException.unrecognizedBiffVersion);
      }

      if (bof.isWorksheet())
      {
        // Read the sheet in
        SheetImpl s = new SheetImpl(excelFile,
                                    sharedStrings,
                                    formattingRecords,
                                    bof,
                                    workbookBof,
                                    nineteenFour,
                                    this);

        BoundsheetRecord br = boundsheets.get
                  (getNumberOfSheets());
        s.setName(br.getName());
        s.setHidden(br.isHidden());
        addSheet(s);
      }
      else if (bof.isChart())
      {
        // Read the sheet in
        SheetImpl s = new SheetImpl(excelFile,
                                    sharedStrings,
                                    formattingRecords,
                                    bof,
                                    workbookBof,
                                    nineteenFour,
                                    this);

        BoundsheetRecord br = boundsheets.get
                  (getNumberOfSheets());
        s.setName(br.getName());
        s.setHidden(br.isHidden());
        addSheet(s);
      }
      else
      {
        LOGGER.warn("BOF is unrecognized");


        while (excelFile.hasNext() && r.getType() != Type.EOF)
        {
          r = excelFile.next();
        }
      }

      // The next record will normally be a BOF or empty padding until
      // the end of the block is reached.  In exceptionally unlucky cases,
      // the last EOF  will coincide with a block division, so we have to
      // check there is more data to retrieve.
      // Thanks to liamg for spotting this
      bof = null;
      if (excelFile.hasNext())
      {
        r = excelFile.next();

        if (r.getType() == Type.BOF)
        {
          bof = new BOFRecord(r);
        }
      }
    }

    // Add all the local names to the specific sheets
    for (NameRecord nr : localNames)
      if (nr.getBuiltInName() == null)
      {
        LOGGER.warn("Usage of a local non-builtin name: " + nr.getName());
      }
      else if (nr.getBuiltInName() == BuiltInName.PRINT_AREA ||
              nr.getBuiltInName() == BuiltInName.PRINT_TITLES)
      {
        // appears to use the internal tab number rather than the
        // external sheet index
        SheetImpl s = sheets.get(nr.getSheetRef() - 1);
        s.addLocalName(nr);
      }
  }

  /**
   * Accessor for the formattingRecords, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the formatting records
   */
  public FormattingRecords getFormattingRecords()
  {
    return formattingRecords;
  }

  /**
   * Accessor for the externSheet, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the external sheet record
   */
  public ExternalSheetRecord getExternalSheetRecord()
  {
    return externSheet;
  }

  /**
   * Accessor for the MsoDrawingGroup, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the Mso Drawing Group record
   */
  public MsoDrawingGroupRecord getMsoDrawingGroupRecord()
  {
    return msoDrawingGroup;
  }

  /**
   * Accessor for the supbook records, used by the WritableWorkbook
   * when creating a copy of this
   *
   * @return the supbook records
   */
  public List<SupbookRecord> getSupbookRecords()
  {
    return Collections.unmodifiableList(supbooks);
  }

  /**
   * Accessor for the name records.  Used by the WritableWorkbook when
   * creating a copy of this
   *
   * @return the array of names
   */
  public List<NameRecord> getNameRecords()
  {
    return Collections.unmodifiableList(nameTable);
  }

  /**
   * Accessor for the fonts, used by the WritableWorkbook
   * when creating a copy of this
   * @return the fonts used in this workbook
   */
  public Fonts getFonts()
  {
    return fonts;
  }

  /**
   * Returns the cell for the specified location eg. "Sheet1!A4".
   * This is identical to using the CellReferenceHelper with its
   * associated performance overheads, consequently it should
   * be use sparingly
   *
   * @param loc the cell to retrieve
   * @return the cell at the specified location
   */
  @Override
  public Cell getCell(String loc)
  {
    Sheet s = getSheet(CellReferenceHelper.getSheet(loc));
    return s.getCell(loc);
  }

  /**
   * Gets the named cell from this workbook.  If the name refers to a
   * range of cells, then the cell on the top left is returned.  If
   * the name cannot be found, null is returned
   *
   * @param  name the name of the cell/range to search for
   * @return the cell in the top left of the range if found, NULL
   *         otherwise
   */
  @Override
  public Cell findCellByName(String name)
  {
    NameRecord nr = namedRecords.get(name);

    if (nr == null)
    {
      return null;
    }

    NameRecord.NameRange[] ranges = nr.getRanges();

    // Go and retrieve the first cell in the first range
    Sheet s = getSheet(getExternalSheetIndex(ranges[0].getExternalSheet()));
    int col = ranges[0].getFirstColumn();
    int row = ranges[0].getFirstRow();

    // If the sheet boundaries fall short of the named cell, then return
    // an empty cell to stop an exception being thrown
    if (col > s.getColumns() || row > s.getRows())
    {
      return new EmptyCell(col, row);
    }

    Cell cell = s.getCell(col, row);

    return cell;
  }

  /**
   * Gets the named range from this workbook.  The Range object returns
   * contains all the cells from the top left to the bottom right
   * of the range.
   * If the named range comprises an adjacent range,
   * the Range[] will contain one object; for non-adjacent
   * ranges, it is necessary to return an array of length greater than
   * one.
   * If the named range contains a single cell, the top left and
   * bottom right cell will be the same cell
   *
   * @param name the name to find
   * @return the range of cells
   */
  @Override
  public Range[] findByName(String name)
  {
    NameRecord nr = namedRecords.get(name);

    if (nr == null)
    {
      return null;
    }

    NameRecord.NameRange[] ranges = nr.getRanges();

    Range[] cellRanges = new Range[ranges.length];

    for (int i = 0; i < ranges.length; i++)
    {
      cellRanges[i] = new RangeImpl
        (this,
         getExternalSheetIndex(ranges[i].getExternalSheet()),
         ranges[i].getFirstColumn(),
         ranges[i].getFirstRow(),
         getLastExternalSheetIndex(ranges[i].getExternalSheet()),
         ranges[i].getLastColumn(),
         ranges[i].getLastRow());
    }

    return cellRanges;
  }

  /**
   * Gets the named ranges
   *
   * @return the list of named cells within the workbook
   */
  @Override
  public String[] getRangeNames()
  {
    return namedRecords.keySet().toArray(new String[namedRecords.size()]);
  }

  /**
   * Method used when parsing formulas to make sure we are trying
   * to parse a supported biff version
   *
   * @return the BOF record
   */
  @Override
  public BOFRecord getWorkbookBof()
  {
    return workbookBof;
  }

  /**
   * Determines whether the sheet is protected
   *
   * @return whether or not the sheet is protected
   */
  @Override
  public boolean isProtected()
  {
    return wbProtected;
  }

  /**
   * Accessor for the settings
   *
   * @return the workbook settings
   */
  public WorkbookSettings getSettings()
  {
    return settings;
  }

  /**
   * Accessor/implementation method for the external sheet reference
   *
   * @param sheetName the sheet name to look for
   * @return the external sheet index
   */
  @Override
  public int getExternalSheetIndex(String sheetName)
  {
    return 0;
  }

  /**
   * Accessor/implementation method for the external sheet reference
   *
   * @param sheetName the sheet name to look for
   * @return the external sheet index
   */
  @Override
  public int getLastExternalSheetIndex(String sheetName)
  {
    return 0;
  }

  /**
   * Gets the name at the specified index
   *
   * @param index the index into the name table
   * @return the name of the cell
   * @exception NameRangeException
   */
  @Override
  public String getName(int index) throws NameRangeException
  {
    if (index < 0 || index >= nameTable.size())
      throw new NameRangeException();

    return nameTable.get(index).getName();
  }

  /**
   * Gets the index of the name record for the name
   *
   * @param name the name to search for
   * @return the index in the name table
   */
  @Override
  public int getNameIndex(String name)
  {
    NameRecord nr = namedRecords.get(name);

    return nr != null ? nr.getIndex() : 0;
  }

  /**
   * Accessor for the drawing group
   *
   * @return  the drawing group
   */
  public DrawingGroup getDrawingGroup()
  {
    return drawingGroup;
  }

  /**
   * Accessor for the CompoundFile.  For this feature to return non-null
   * value, the propertySets feature in WorkbookSettings must be enabled
   * and the workbook must contain additional property sets.  This
   * method is used during the workbook copy
   *
   * @return the base compound file if it contains additional data items
   *         and property sets are enabled
   */
  public CompoundFile getCompoundFile()
  {
    return excelFile.getCompoundFile();
  }

  /**
   * Accessor for the containsMacros
   *
   * @return TRUE if this workbook contains macros, FALSE otherwise
   */
  public boolean containsMacros()
  {
    return containsMacros;
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
   * Accessor for the country record, using during copying
   *
   * @return the country record read in
   */
  public CountryRecord getCountryRecord()
  {
    return countryRecord;
  }

  /**
   * Accessor for addin function names
   *
   * @return immutable list of add in function names
   */
  public List<String> getAddInFunctionNames()
  {
    return Collections.unmodifiableList(addInFunctions);
  }

  /**
   * Gets the sheet index in this workbook.  Used when importing a sheet
   *
   * @param sheet the sheet
   * @return the 0-based sheet index, or -1 if it is not found
   */
  public int getIndex(Sheet sheet)
  {
    String name = sheet.getName();
    int index = -1;
    int pos = 0;

    for (Iterator<BoundsheetRecord> i = boundsheets.iterator() ; i.hasNext() && index == -1 ;)
    {
      BoundsheetRecord br = i.next();

      if (br.getName().equals(name))
      {
        index = pos;
      }
      else
      {
        pos++;
      }
    }

    return index;
  }

  /**
   *
   * @return immutable list af XCTRecords
   */
  public List<XCTRecord> getXCTRecords()
  {
    return Collections.unmodifiableList(xctRecords);
  }

}