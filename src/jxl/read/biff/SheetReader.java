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
import jxl.*;
import jxl.CellReferenceHelper;
import jxl.HeaderFooter;
import jxl.biff.*;
import jxl.biff.drawing.*;
import jxl.biff.formula.FormulaException;
import jxl.common.*;
import jxl.format.*;

/**
 * Reads the sheet.  This functionality was originally part of the
 * SheetImpl class, but was separated out in order to simplify the former
 * class
 */
final class SheetReader
{
  /**
   * The logger
   */
  private static final Logger LOGGER = Logger.getLogger(SheetReader.class);

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
  private final int startPosition;

  /**
   * The list of non-default row properties
   */
  private final List<RowRecord> rowProperties = new ArrayList<>(10);

  /**
   * An array of column info records.  They are held this way before
   * they are transferred to the more convenient array
   */
  private final List<ColumnInfoRecord> columnInfosArray = new ArrayList<>();

  /**
   * A list of shared formula groups
   */
  private final List<SharedFormulaRecord> sharedFormulas = new ArrayList<>();

  /**
   * A list of hyperlinks on this page
   */
  private final List<Hyperlink> hyperlinks = new ArrayList<>();

  /**
   * The list of conditional formats on this page
   */
  private final List<ConditionalFormat> conditionalFormats = new ArrayList<>();

  /**
   * The autofilter information
   */
  private AutoFilter autoFilter;

  /**
   * A list of merged cells on this page
   */
  private final List<Range> mergedCells = new ArrayList<>();

  /**
   * The list of data validations on this page
   */
  private DataValidation dataValidation;

  /**
   * The list of charts on this page
   */
  private final List<Chart> charts = new ArrayList<>();

  /**
   * The list of drawings on this page
   */
  private final List<DrawingGroupObject> drawings = new ArrayList<>();

  /**
   * The drawing data for the drawings
   */
  private DrawingData drawingData;

  /**
   * Indicates whether or not the dates are based around the 1904 date system
   */
  private final boolean nineteenFour;

  /**
   * The PLS print record
   */
  private PLSRecord plsRecord;

  /**
   * The property set record associated with this workbook
   */
  private ButtonPropertySetRecord buttonPropertySet;

  /**
   * The workspace options
   */
  private WorkspaceInformationRecord workspaceOptions;

  /**
   * The horizontal page breaks contained on this sheet
   */
  private HorizontalPageBreaksRecord rowBreaks = new HorizontalPageBreaksRecord();

  /**
   * The vertical page breaks contained on this sheet
   */
  private VerticalPageBreaksRecord columnBreaks = new VerticalPageBreaksRecord();

  /**
   * The maximum row outline level
   */
  private int maxRowOutlineLevel;

  /**
   * The maximum column outline level
   */
  private int maxColumnOutlineLevel;

  /**
   * The sheet settings
   */
  private final SheetSettings settings;

  /**
   * The workbook settings
   */
  private final WorkbookSettings workbookSettings;

  /**
   * A handle to the workbook which contains this sheet.  Some of the records
   * need this in order to reference external sheets
   */
  private final WorkbookParser workbook;

  /**
   * A handle to the sheet
   */
  private final SheetImpl sheet;

  /**
   * Constructor
   *
   * @param fr the formatting records
   * @param sst the shared string table
   * @param f the excel file
   * @param sb the bof record which indicates the start of the sheet
   * @param wb the bof record which indicates the start of the sheet
   * @param wp the workbook which this sheet belongs to
   * @param sp the start position of the sheet bof in the excel file
   * @param sh the sheet
   * @param nf 1904 date record flag
   * @exception BiffException
   */
  SheetReader(File f,
              SSTRecord sst,
              FormattingRecords fr,
              BOFRecord sb,
              BOFRecord wb,
              boolean nf,
              WorkbookParser wp,
              int sp,
              SheetImpl sh)
  {
    excelFile = f;
    sharedStrings = sst;
    formattingRecords = fr;
    sheetBof = sb;
    workbookBof = wb;
    nineteenFour = nf;
    workbook = wp;
    startPosition = sp;
    sheet = sh;
    settings = new SheetSettings(sh);
    workbookSettings = workbook.getSettings();
  }

  /**
   * Adds the cell to the array
   *
   * @param cell the cell to add
   */
  private void addCell(Cell cell)
  {
    jxl.CellCoordinate coord = cell.getCoordinate();
    if (cells.put(coord, cell) != null)
    {
      StringBuffer sb = new StringBuffer();
      CellReferenceHelper.getCellReference(cell.getColumn(), cell.getRow(), sb);
      LOGGER.warn("Cell " + sb.toString() + " already contains data");
    }
    numRows = Math.max(numRows, cell.getRow() + 1);
    numCols = Math.max(numCols, cell.getColumn() + 1);
  }

  /**
   * Reads in the contents of this sheet
   */
  final void read()
  {
    BaseSharedFormulaRecord sharedFormula = null;
    boolean sharedFormulaAdded = false;

    boolean cont = true;

    // Set the position within the file
    excelFile.setPos(startPosition);

    // Handles to the last drawing and obj records
    MsoDrawingRecord msoRecord = null;
    ObjRecord objRecord = null;
    boolean firstMsoRecord = true;

    // Handle to the last conditional format record
    ConditionalFormat condFormat = null;

    // Handle to the autofilter records
    FilterModeRecord filterMode = null;
    AutoFilterInfoRecord autoFilterInfo = null;

    // A handle to window2 record
    Window2Record window2Record = null;

    // Hash map of comments, indexed on objectId.  As each corresponding
    // note record is encountered, these are removed from the array
    HashMap<Integer, Comment> comments = new HashMap<>();

    // A list of object ids - used for cross referencing
    ArrayList<Integer> objectIds = new ArrayList<>();

    // A handle to a continue record read in
    ContinueRecord continueRecord = null;

    boolean first = true;

    while (cont)
    {
      Record r = excelFile.next();
      Type type = r.getType();

      if (type == Type.UNKNOWN && r.getCode() == 0)
      {
        LOGGER.warn("Biff code zero found");

        // Try a dimension record
        if (r.getLength() == 0xa)
        {
          LOGGER.warn("Biff code zero found - trying a dimension record.");
          r.setType(Type.DIMENSION);
        }
        else
        {
          LOGGER.warn("Biff code zero found - Ignoring.");
        }
      }

      if (first && type != Type.DIMENSION) {
        numRows = workbookSettings.getStartRowCount();
        numCols = workbookSettings.getStartColumnCount();
        first = false;
      }

      switch (type) {
        case DIMENSION:
          DimensionRecord dr = workbookBof.isBiff8()
                  ? new DimensionRecord(r)
                  : new DimensionRecord(r, DimensionRecord.biff7);
          numRows = dr.getNumberOfRows();
          numCols = dr.getNumberOfColumns();
          break;

        case LABELSST:
          addCell(new LabelSSTRecord(r,
                  sharedStrings,
                  formattingRecords,
                  sheet));
          break;

        case RK:
        case RK2:
          RKRecord rkr = new RKRecord(r, formattingRecords, sheet);
          if (formattingRecords.isDate(rkr.getXFIndex())) {
            DateCell dc = new DateRecord(rkr, rkr.getXFIndex(), formattingRecords, nineteenFour, sheet);
            addCell(dc);
          } else
            addCell(rkr);
          break;

        case HLINK: {
          HyperlinkRecord hr = new HyperlinkRecord(r, sheet, workbookSettings);
          hyperlinks.add(hr);
          break;
        }

        case MERGEDCELLS:
          MergedCellsRecord mc = new MergedCellsRecord(r, sheet);
          mergedCells.addAll(mc.getRanges());
          break;

        case MULRK: {
          MulRKRecord mulrk = new MulRKRecord(r);
          // Get the individual cell records from the multiple record
          int num = mulrk.getNumberOfColumns();
          for (int i = 0; i < num; i++) {
            int ixf = mulrk.getXFIndex(i);

            NumberValue nv = new NumberValue(mulrk.getRow(),
                    mulrk.getFirstColumn() + i,
                    RKHelper.getDouble(mulrk.getRKNumber(i)),
                    ixf,
                    formattingRecords,
                    sheet);

            if (formattingRecords.isDate(ixf)) {
              DateCell dc = new DateRecord(nv,
                      ixf,
                      formattingRecords,
                      nineteenFour,
                      sheet);
              addCell(dc);
            } else {
              nv.setNumberFormat(formattingRecords.getNumberFormat(ixf));
              addCell(nv);
            }
          }
        }
        break;

        case NUMBER: {
          NumberRecord nr = new NumberRecord(r, formattingRecords, sheet);
          if (formattingRecords.isDate(nr.getXFIndex())) {
            DateCell dc = new DateRecord(nr,
                    nr.getXFIndex(),
                    formattingRecords,
                    nineteenFour, sheet);
            addCell(dc);
          } else
            addCell(nr);
        }
        break;

        case BOOLERR: {
          BooleanRecord br = new BooleanRecord(r, formattingRecords, sheet);
          if (br.isError()) {
            ErrorRecord er = new ErrorRecord(br.getRecord(), formattingRecords,
                    sheet);
            addCell(er);
          } else
            addCell(br);
          break;
        }

        case PRINTGRIDLINES:
          // A handle to printgridlines record
          PrintGridLinesRecord printGridLinesRecord = new PrintGridLinesRecord(r);
          settings.setPrintGridLines(printGridLinesRecord.getPrintGridLines());
          break;

        case PRINTHEADERS:
          // A handle to printheaders record
          PrintHeadersRecord printHeadersRecord = new PrintHeadersRecord(r);
          settings.setPrintHeaders(printHeadersRecord.getPrintHeaders());
          break;

        case WINDOW2:
          window2Record = workbookBof.isBiff8()
                  ? new Window2Record(r)
                  : new Window2Record(r, Window2Record.biff7);
          settings.setShowGridLines(window2Record.getShowGridLines());
          settings.setDisplayZeroValues(window2Record.getDisplayZeroValues());
          settings.setSelected(true);
          settings.setPageBreakPreviewMode(window2Record.isPageBreakPreview());
          break;

        case PANE: {
          PaneRecord pr = new PaneRecord(r);
          if (window2Record != null
                  && window2Record.getFrozen()) {
            settings.setVerticalFreeze(pr.getRowsVisible());
            settings.setHorizontalFreeze(pr.getColumnsVisible());
          }
          break;
        }

        case CONTINUE:
          // don't know what this is for, but keep hold of it anyway
          continueRecord = new ContinueRecord(r);
          break;

        case NOTE:
          if (!workbookSettings.getDrawingsDisabled()) {
            NoteRecord nr = new NoteRecord(r);

            // Get the comment for the object id
            Comment comment = comments.remove(nr.getObjectId());

            if (comment == null)
              LOGGER.warn(" cannot find comment for note id "
                      + nr.getObjectId() + "...ignoring");
            else {
              comment.setNote(nr);

              drawings.add(comment);

              addCellComment(comment.getColumn(),
                      comment.getRow(),
                      comment.getText(),
                      comment.getWidth(),
                      comment.getHeight());
            }
          }
          break;

        case ARRAY:
          break;

        case PROTECT: {
          ProtectRecord pr = new ProtectRecord(r);
          settings.setProtected(pr.isProtected());
          break;
        }

        case SHAREDFORMULA:
          if (sharedFormula == null) {
            LOGGER.warn("Shared template formula is null - "
                    + "trying most recent formula template");
            SharedFormulaRecord lastSharedFormula
                    = sharedFormulas.get(sharedFormulas.size() - 1);

            if (lastSharedFormula != null)
              sharedFormula = lastSharedFormula.getTemplateFormula();
          }
          SharedFormulaRecord sfr = new SharedFormulaRecord(r, sharedFormula, workbook, workbook, sheet);
          sharedFormulas.add(sfr);
          sharedFormula = null;
          break;

        case FORMULA:
        case FORMULA2: {
          FormulaRecord fr = new FormulaRecord(r,
                  excelFile,
                  formattingRecords,
                  workbook,
                  workbook,
                  sheet,
                  workbookSettings,
                  workbook.getWorkbookBof().isBiff8());
          if (fr.isShared()) {
            BaseSharedFormulaRecord prevSharedFormula = sharedFormula;
            sharedFormula = (BaseSharedFormulaRecord) fr.getFormula();

            // See if it fits in any of the shared formulas
            sharedFormulaAdded = addToSharedFormulas(sharedFormula);

            if (sharedFormulaAdded)
              sharedFormula = prevSharedFormula;

            // If we still haven't added the previous base shared formula,
            // revert it to an ordinary formula and add it to the cell
            if (!sharedFormulaAdded && prevSharedFormula != null)
              // Do nothing.  It's possible for the biff file to contain the
              // record sequence
              // FORMULA-SHRFMLA-FORMULA-SHRFMLA-FORMULA-FORMULA-FORMULA
              // ie. it first lists all the formula templates, then it
              // lists all the individual formulas
              addCell(revertSharedFormula(prevSharedFormula));
          } else {
            Cell cell = fr.getFormula();
            try {
              // See if the formula evaluates to date
              if (fr.getFormula().getType() == CellType.NUMBER_FORMULA) {
                NumberFormulaRecord nfr = (NumberFormulaRecord) fr.getFormula();
                if (formattingRecords.isDate(nfr.getXFIndex()))
                  cell = new DateFormulaRecord(nfr,
                          formattingRecords,
                          workbook,
                          workbook,
                          nineteenFour,
                          sheet);
              }

              addCell(cell);
            } catch (FormulaException e) {
              // Something has gone wrong trying to read the formula data eg. it
              // might be unsupported biff7 data
              LOGGER.warn(CellReferenceHelper
                      .getCellReference(cell.getColumn(), cell.getRow()) + " " + e
                      .getMessage());
            }
          }
          break;
        }

        case LABEL:
          addCell(workbookBof.isBiff8()
                  ? new LabelRecord(r, formattingRecords, sheet)
                  : new LabelRecord(r, formattingRecords, sheet, workbookSettings,
                          LabelRecord.biff7));
          break;

        case RSTRING:
          // RString records are obsolete in biff 8
          Assert.verify(!workbookBof.isBiff8());
          RStringRecord lr = new RStringRecord(r, formattingRecords,
                  sheet, workbookSettings,
                  RStringRecord.biff7);
          addCell(lr);
          break;

        case NAME:
          break;

        case PASSWORD: {
          PasswordRecord pr = new PasswordRecord(r);
          settings.setPasswordHash(pr.getPasswordHash());
          break;
        }

        case ROW:
          RowRecord rr = new RowRecord(r);
          // See if the row has anything funny about it
          if (!rr.isDefaultHeight()
                  || !rr.matchesDefaultFontHeight()
                  || rr.isCollapsed()
                  || rr.hasDefaultFormat()
                  || rr.getOutlineLevel() != 0)
            rowProperties.add(rr);
          break;

        case BLANK:
          if (!workbookSettings.getIgnoreBlanks())
            addCell(new BlankCell(r, formattingRecords, sheet));
          break;

        case MULBLANK:
          if (!workbookSettings.getIgnoreBlanks()) {
            MulBlankRecord mulblank = new MulBlankRecord(r);

            // Get the individual cell records from the multiple record
            int num = mulblank.getNumberOfColumns();

            for (int i = 0; i < num; i++) {
              int ixf = mulblank.getXFIndex(i);

              MulBlankCell mbc = new MulBlankCell(mulblank.getRow(),
                      mulblank.getFirstColumn() + i,
                      ixf,
                      formattingRecords,
                      sheet);

              addCell(mbc);
            }
          }
          break;

        case SCL:
          SCLRecord scl = new SCLRecord(r);
          settings.setZoomFactor(scl.getZoomFactor());
          break;

        case COLINFO:
          ColumnInfoRecord cir = new ColumnInfoRecord(r);
          columnInfosArray.add(cir);
          break;

        case HEADER: {
          HeaderRecord hr = workbookBof.isBiff8()
                  ? new HeaderRecord(r)
                  : new HeaderRecord(r, workbookSettings, HeaderRecord.biff7);
          HeaderFooter header = new HeaderFooter(hr.getHeader());
          settings.setHeader(header);
          break;
        }

        case FOOTER: {
          FooterRecord fr = workbookBof.isBiff8()
                  ? new FooterRecord(r)
                  : new FooterRecord(r, workbookSettings, FooterRecord.biff7);
          HeaderFooter footer = new HeaderFooter(fr.getFooter());
          settings.setFooter(footer);
          break;
        }

        case SETUP:
          SetupRecord sr = new SetupRecord(r);
          // If the setup record has its not initialized bit set, then
          // use the sheet settings default values
          if (sr.getInitialized()) {
            if (sr.isPortrait())
              settings.setOrientation(PageOrientation.PORTRAIT);
            else
              settings.setOrientation(PageOrientation.LANDSCAPE);
            if (sr.isRightDown())
              settings.setPageOrder(PageOrder.RIGHT_THEN_DOWN);
            else
              settings.setPageOrder(PageOrder.DOWN_THEN_RIGHT);
            settings.setPaperSize(PaperSize.getPaperSize(sr.getPaperSize()));
            settings.setHeaderMargin(sr.getHeaderMargin());
            settings.setFooterMargin(sr.getFooterMargin());
            settings.setScaleFactor(sr.getScaleFactor());
            settings.setPageStart(sr.getPageStart());
            settings.setFitWidth(sr.getFitWidth());
            settings.setFitHeight(sr.getFitHeight());
            settings.setHorizontalPrintResolution(sr
                    .getHorizontalPrintResolution());
            settings.setVerticalPrintResolution(sr.getVerticalPrintResolution());
            settings.setCopies(sr.getCopies());

            if (workspaceOptions != null)
              settings.setFitToPages(workspaceOptions.getFitToPages());
          }
          break;

        case WSBOOL:
          workspaceOptions = new WorkspaceInformationRecord(r);
          break;

        case DEFCOLWIDTH:
          DefaultColumnWidthRecord dcwr = new DefaultColumnWidthRecord(r);
          settings.setDefaultColumnWidth(dcwr.getWidth());
          break;

        case DEFAULTROWHEIGHT:
          DefaultRowHeightRecord drhr = new DefaultRowHeightRecord(r);
          if (drhr.getHeight() != 0)
            settings.setDefaultRowHeight(drhr.getHeight());
          break;

        case CONDFMT:
          ConditionalFormatRangeRecord cfrr
                  = new ConditionalFormatRangeRecord(r);
          condFormat = new ConditionalFormat(cfrr);
          conditionalFormats.add(condFormat);
          break;

        case CF:
          ConditionalFormatRecord cfr = new ConditionalFormatRecord(r);
          condFormat.addCondition(cfr);
          break;

        case FILTERMODE:
          filterMode = new FilterModeRecord(r);
          break;

        case AUTOFILTERINFO:
          autoFilterInfo = new AutoFilterInfoRecord(r);
          break;

        case AUTOFILTER:
          if (!workbookSettings.getAutoFilterDisabled()) {
            AutoFilterRecord af = new AutoFilterRecord(r);

            if (autoFilter == null) {
              autoFilter = new AutoFilter(filterMode, autoFilterInfo);
              filterMode = null;
              autoFilterInfo = null;
            }

            autoFilter.add(af);
          }
          break;

        case LEFTMARGIN: {
          MarginRecord m = new LeftMarginRecord(r);
          settings.setLeftMargin(m.getMargin());
          break;
        }

        case RIGHTMARGIN: {
          MarginRecord m = new RightMarginRecord(r);
          settings.setRightMargin(m.getMargin());
          break;
        }

        case TOPMARGIN: {
          MarginRecord m = new TopMarginRecord(r);
          settings.setTopMargin(m.getMargin());
          break;
        }

        case BOTTOMMARGIN: {
          MarginRecord m = new BottomMarginRecord(r);
          settings.setBottomMargin(m.getMargin());
          break;
        }

        case HORIZONTALPAGEBREAKS:
          rowBreaks = workbookBof.isBiff8()
                  ? new HorizontalPageBreaksRecord(r)
                  : new HorizontalPageBreaksRecord(r, HorizontalPageBreaksRecord.biff7);
          break;

        case VERTICALPAGEBREAKS:
          columnBreaks = workbookBof.isBiff8()
                  ? new VerticalPageBreaksRecord(r)
                  : new VerticalPageBreaksRecord(r, VerticalPageBreaksRecord.biff7);
          break;

        case PLS:
          plsRecord = new PLSRecord(r);
          // Check for Continue records
          while (excelFile.peek().getType() == Type.CONTINUE)
            r.addContinueRecord(excelFile.next());
          break;

        case DVAL:
          if (!workbookSettings.getCellValidationDisabled()) {
            DataValidityListRecord dvlr = new DataValidityListRecord(r);
            if (dvlr.getObjectId() == -1)
              if (msoRecord != null && objRecord == null) {
                // there is a drop down associated with this data validation
                if (drawingData == null)
                  drawingData = new DrawingData();

                Drawing2 d2 = new Drawing2(msoRecord, drawingData,
                        workbook.getDrawingGroup());
                drawings.add(d2);
                msoRecord = null;

                dataValidation = new DataValidation(dvlr);
              } else
                // no drop down
                dataValidation = new DataValidation(dvlr);
            else if (objectIds.contains(dvlr.getObjectId()))
              dataValidation = new DataValidation(dvlr);
            else
              LOGGER.warn("object id " + dvlr.getObjectId() + " referenced "
                      + " by data validity list record not found - ignoring");
          }
          break;

        case HCENTER: {
          CentreRecord hr = new CentreRecord(r);
          settings.setHorizontalCentre(hr.isCentre());
          break;
        }

        case VCENTER:
          CentreRecord vc = new CentreRecord(r);
          settings.setVerticalCentre(vc.isCentre());
          break;

        case DV:
          if (!workbookSettings.getCellValidationDisabled()) {
            DataValiditySettingsRecord dvsr
                    = new DataValiditySettingsRecord(r,
                            workbook,
                            workbook,
                            workbook.getSettings());
            if (dataValidation != null) {
              dataValidation.add(dvsr);
              addCellValidation(dvsr.getFirstColumn(),
                      dvsr.getFirstRow(),
                      dvsr.getLastColumn(),
                      dvsr.getLastRow(),
                      dvsr);
            } else
              LOGGER.warn("cannot add data validity settings");
          }
          break;

        case OBJ:
          objRecord = new ObjRecord(r);
          if (!workbookSettings.getDrawingsDisabled()) {
            // sometimes excel writes out continue records instead of drawing
            // records, so forcibly hack the stashed continue record into
            // a drawing record
            if (msoRecord == null && continueRecord != null) {
              LOGGER.warn("Cannot find drawing record - using continue record");
              msoRecord = new MsoDrawingRecord(continueRecord.getRecord());
              continueRecord = null;
            }
            handleObjectRecord(objRecord, msoRecord, comments);
            objectIds.add(objRecord.getObjectId());
          } // Save chart handling until the chart BOF record appears
          if (objRecord.getType() != ObjRecord.CHART) {
            objRecord = null;
            msoRecord = null;
          }
          break;

        case MSODRAWING:
          if (!workbookSettings.getDrawingsDisabled()) {
            if (msoRecord != null)
              // For form controls, a rogue MSODRAWING record can crop up
              // after the main one.  Add these into the drawing data
              drawingData.addRawData(msoRecord.getData());
            msoRecord = new MsoDrawingRecord(r);

            if (firstMsoRecord) {
              msoRecord.setFirst();
              firstMsoRecord = false;
            }
          }
          break;

        case BUTTONPROPERTYSET:
          buttonPropertySet = new ButtonPropertySetRecord(r);
          break;

        case CALCMODE: {
          CalcModeRecord cmr = new CalcModeRecord(r);
          settings.setAutomaticFormulaCalculation(cmr.isAutomatic());
          break;
        }

        case SAVERECALC: {
          SaveRecalcRecord cmr = new SaveRecalcRecord(r);
          settings.setRecalculateFormulasBeforeSave(cmr.getRecalculateOnSave());
          break;
        }

        case GUTS:
          GuttersRecord gr = new GuttersRecord(r);
          maxRowOutlineLevel
                  = gr.getRowOutlineLevel() > 0 ? gr.getRowOutlineLevel() - 1 : 0;
          maxColumnOutlineLevel
                  = gr.getColumnOutlineLevel() > 0 ? gr.getRowOutlineLevel() - 1 : 0;
          break;

        case BOF: {
          BOFRecord br = new BOFRecord(r);
          Assert.verify(!br.isWorksheet());
          int startpos = excelFile.getPos() - r.getLength() - 4;
          // Skip to the end of the nested bof
          // Thanks to Rohit for spotting this
          Record r2 = excelFile.next();
          while (r2.getCode() != Type.EOF.value)
            r2 = excelFile.next();
          if (br.isChart()) {
            if (!workbook.getWorkbookBof().isBiff8())
              LOGGER.warn("only biff8 charts are supported");
            else {
              if (drawingData == null)
                drawingData = new DrawingData();

              if (!workbookSettings.getDrawingsDisabled()) {
                Chart chart = new Chart(msoRecord, objRecord, drawingData,
                        startpos, excelFile.getPos(),
                        excelFile, workbookSettings);
                charts.add(chart);

                if (workbook.getDrawingGroup() != null)
                  workbook.getDrawingGroup().add(chart);
              }
            }

            // Reset the drawing records
            msoRecord = null;
            objRecord = null;
          }   // If this worksheet is just a chart, then the EOF reached
          // represents the end of the sheet as well as the end of the chart
          if (sheetBof.isChart())
            cont = false;
          break;
        }

        case EOF:
          cont = false;
          break;
      }
    }

    // Restore the file to its accurate position
    excelFile.restorePos();

    // Add all the shared formulas to the sheet as individual formulas
    for (SharedFormulaRecord sfr : sharedFormulas) {
      Cell[] sfnr = sfr.getFormulas(formattingRecords, nineteenFour);

      for (Cell cell : sfnr)
        addCell(cell);
    }

    // If the last base shared formula wasn't added to the sheet, then
    // revert it to an ordinary formula and add it
    if (!sharedFormulaAdded && sharedFormula != null)
    {
      addCell(revertSharedFormula(sharedFormula));
    }

    // If there is a stray msoDrawing record, then flag to the drawing group
    // that one has been omitted
    if (msoRecord != null && workbook.getDrawingGroup() != null)
    {
      workbook.getDrawingGroup().setDrawingsOmitted(msoRecord, objRecord);
    }

    // Check that the comments hash is empty
    if (!comments.isEmpty())
    {
      LOGGER.warn("Not all comments have a corresponding Note record");
    }
  }

  /**
   * Sees if the shared formula belongs to any of the shared formula
   * groups
   *
   * @param fr the candidate shared formula
   * @return TRUE if the formula was added, FALSE otherwise
   */
  private boolean addToSharedFormulas(BaseSharedFormulaRecord fr)
  {
    boolean added = false;

    for (int i=0, size=sharedFormulas.size(); i<size && !added; ++i)
      added = sharedFormulas.get(i).add(fr);

    return added;
  }

  /**
   * Reverts the shared formula passed in to an ordinary formula and adds
   * it to the list
   *
   * @param f the formula
   * @return the new formula
   * @exception FormulaException
   */
  private Cell revertSharedFormula(BaseSharedFormulaRecord f)
  {
    // String formulas look for a STRING record soon after the formula
    // occurred.  Temporarily the position in the excel file back
    // to the point immediately after the formula record
    int pos = excelFile.getPos();
    excelFile.setPos(f.getFilePos());

    FormulaRecord fr = new FormulaRecord(f.getRecord(),
                                         excelFile,
                                         formattingRecords,
                                         workbook,
                                         workbook,
                                         FormulaRecord.ignoreSharedFormula,
                                         sheet,
                                         workbookSettings,
                                         workbook.getWorkbookBof().isBiff8());

    try
    {
    Cell cell = fr.getFormula();

    // See if the formula evaluates to date
    if (fr.getFormula().getType() == CellType.NUMBER_FORMULA)
    {
      NumberFormulaRecord nfr = (NumberFormulaRecord) fr.getFormula();
      if (formattingRecords.isDate(fr.getXFIndex()))
      {
        cell = new DateFormulaRecord(nfr,
                                     formattingRecords,
                                     workbook,
                                     workbook,
                                     nineteenFour,
                                     sheet);
      }
    }

    excelFile.setPos(pos);
    return cell;
    }
    catch (FormulaException e)
    {
      // Something has gone wrong trying to read the formula data eg. it
      // might be unsupported biff7 data
      LOGGER.warn
        (CellReferenceHelper.getCellReference(fr.getColumn(), fr.getRow()) +
         " " + e.getMessage());

      return null;
    }
  }


  /**
   * Accessor
   *
   * @return the number of rows
   */
  final int getNumRows()
  {
    return numRows;
  }

  /**
   * Accessor
   *
   * @return the number of columns
   */
  final int getNumCols()
  {
    return numCols;
  }

  /**
   * Accessor
   *
   * @return the cells
   */
  final Map<CellCoordinate, Cell> getCells()
  {
    return Collections.unmodifiableMap(cells);
  }

  /**
   * Accessor
   *
   * @return the row properties
   */
  final List<RowRecord> getRowProperties()
  {
    return rowProperties;
  }

  /**
   * Accessor
   *
   * @return the column information
   */
  final List<ColumnInfoRecord> getColumnInfosArray()
  {
    return columnInfosArray;
  }

  /**
   * Accessor
   *
   * @return the hyperlinks
   */
  final List<Hyperlink> getHyperlinks()
  {
    return hyperlinks;
  }

  /**
   * Accessor
   *
   * @return the conditional formatting
   */
  final List<ConditionalFormat> getConditionalFormats()
  {
    return conditionalFormats;
  }

  /**
   * Accessor
   *
   * @return the autofilter
   */
  final AutoFilter getAutoFilter()
  {
    return autoFilter;
  }

  /**
   * Accessor
   *
   * @return the charts
   */
  final List<Chart> getCharts()
  {
    return charts;
  }

  /**
   * Accessor
   *
   * @return the drawings
   */
  final List<DrawingGroupObject> getDrawings()
  {
    return drawings;
  }

  /**
   * Accessor
   *
   * @return the data validations
   */
  final DataValidation getDataValidation()
  {
    return dataValidation;
  }

  /**
   * Accessor
   *
   * @return the ranges
   */
  final List<Range> getMergedCells()
  {
    return mergedCells;
  }

  /**
   * Accessor
   *
   * @return the sheet settings
   */
  final SheetSettings getSettings()
  {
    return settings;
  }

  /**
   * Accessor
   *
   * @return the row breaks
   */
  final HorizontalPageBreaksRecord getRowBreaks()
  {
    return rowBreaks;
  }

  /**
   * Accessor
   *
   * @return the column breaks
   */
  final VerticalPageBreaksRecord getColumnBreaks()
  {
    return columnBreaks;
  }

  /**
   * Accessor
   *
   * @return the workspace options
   */
  final WorkspaceInformationRecord getWorkspaceOptions()
  {
    return workspaceOptions;
  }

  /**
   * Accessor
   *
   * @return the environment specific print record
   */
  final PLSRecord getPLS()
  {
    return plsRecord;
  }

  /**
   * Accessor for the button property set, used during copying
   *
   * @return the button property set
   */
  final ButtonPropertySetRecord getButtonPropertySet()
  {
    return buttonPropertySet;
  }

  /**
   * Adds a cell comment to a cell just read in
   *
   * @param col the column for the comment
   * @param row the row for the comment
   * @param text the comment text
   * @param width the width of the comment text box
   * @param height the height of the comment text box
   */
  private void addCellComment(int col,
                              int row,
                              String text,
                              double width,
                              double height)
  {
    Cell c = cells.computeIfAbsent(
            new CellCoordinate(col, row),
            coord -> new MulBlankCell(
                    coord,
                    0,
                    formattingRecords,
                    sheet));

    if (c instanceof CellFeaturesAccessor)
    {
      CellFeaturesAccessor cv = (CellFeaturesAccessor) c;
      CellFeatures cf = cv.getCellFeatures();

      if (cf == null)
      {
        cf = new CellFeatures();
        cv.setCellFeatures(cf);
      }

      cf.setReadComment(text, width ,height);
    }
    else
    {
      LOGGER.warn("Not able to add comment to cell type " +
                  c.getClass().getName() +
                  " at " + CellReferenceHelper.getCellReference(col, row));
    }
  }

  /**
   * Adds a cell comment to a cell just read in
   *
   * @param col1 the column for the comment
   * @param row1 the row for the comment
   * @param col2 the row for the comment
   * @param row2 the row for the comment
   * @param dvsr the validation settings
   */
  private void addCellValidation(int col1,
                                 int row1,
                                 int col2,
                                 int row2,
                                 DataValiditySettingsRecord dvsr)
  {
    for (int row = row1; row <= row2; row++)
    {
      for (int col = col1; col <= col2; col++)
      {
        Cell c = cells.get(new CellCoordinate(col, row));

        if (c == null) {
          MulBlankCell mbc = new MulBlankCell(row,
                                              col,
                                              0,
                                              formattingRecords,
                                              sheet);
          CellFeatures cf = new CellFeatures();
          cf.setValidationSettings(dvsr);
          mbc.setCellFeatures(cf);
          addCell(mbc);
        } else if (c instanceof CellFeaturesAccessor) {
          // Check to see if the cell already contains a comment
          CellFeaturesAccessor cv = (CellFeaturesAccessor) c;
          CellFeatures cf = cv.getCellFeatures();

          if (cf == null)
          {
            cf = new CellFeatures();
            cv.setCellFeatures(cf);
          }

          cf.setValidationSettings(dvsr);
        }
        else
        {
          LOGGER.warn("Not able to add comment to cell type " +
                      c.getClass().getName() +
                      " at " + CellReferenceHelper.getCellReference(col, row));
        }
      }
    }
  }

  /**
   * Reads in the object record
   *
   * @param objRecord the obj record
   * @param msoRecord the mso drawing record read in earlier
   * @param comments the hash map of comments
   */
  private void handleObjectRecord(ObjRecord objRecord,
                                  MsoDrawingRecord msoRecord,
                                  HashMap<Integer, Comment> comments)
  {
    if (msoRecord == null)
    {
      LOGGER.warn("Object record is not associated with a drawing " +
                  " record - ignoring");
      return;
    }

    try
    {
      // Handle images
      if (objRecord.getType() == ObjRecord.PICTURE)
      {
        if (drawingData == null)
        {
          drawingData = new DrawingData();
        }

        Drawing drawing = new Drawing(msoRecord,
                                      objRecord,
                                      drawingData,
                                      workbook.getDrawingGroup(),
                                      sheet);
        drawings.add(drawing);
        return;
      }

      // Handle comments
      if (objRecord.getType() == ObjRecord.EXCELNOTE)
      {
        if (drawingData == null)
        {
          drawingData = new DrawingData();
        }

        Comment comment = workbookBof.isBiff8()
                ? new CommentBiff8(msoRecord,
                                      objRecord,
                                      drawingData,
                                      workbook.getDrawingGroup())
                : new CommentBiff7(msoRecord,
                                      objRecord,
                                      drawingData,
                                      workbook.getDrawingGroup(),
                                      workbookSettings);

        // Sometimes Excel writes out Continue records instead of drawing
        // records, so forcibly hack all of these into a drawing record
        Record r2 = excelFile.next();
        if (r2.getType() == Type.MSODRAWING || r2.getType() == Type.CONTINUE)
        {
          MsoDrawingRecord mso = new MsoDrawingRecord(r2);
          comment.addMso(mso);
          r2 = excelFile.next();
        }
        Assert.verify(r2.getType() == Type.TXO);
        TextObjectRecord txo = new TextObjectRecord(r2);
        comment.setTextObject(txo);

        r2 = excelFile.next();
        Assert.verify(r2.getType() == Type.CONTINUE);
        ContinueRecord text = new ContinueRecord(r2);
        comment.setText(text);

        r2 = excelFile.next();
        if (r2.getType() == Type.CONTINUE)
        {
          ContinueRecord formatting = new ContinueRecord(r2);
          comment.setFormatting(formatting);
        }

        comments.put(comment.getObjectId(), comment);
        return;
      }

      // Handle combo boxes
      if (objRecord.getType() == ObjRecord.COMBOBOX)
      {
        if (drawingData == null)
        {
          drawingData = new DrawingData();
        }

        ComboBox comboBox = new ComboBox(msoRecord,
                                         objRecord,
                                         drawingData,
                                         workbook.getDrawingGroup(),
                                         workbookSettings);
        drawings.add(comboBox);
        return;
      }

      // Handle check boxes
      if (objRecord.getType() == ObjRecord.CHECKBOX)
      {
        if (drawingData == null)
        {
          drawingData = new DrawingData();
        }

        CheckBox checkBox = new CheckBox(msoRecord,
                                         objRecord,
                                         drawingData,
                                         workbook.getDrawingGroup(),
                                         workbookSettings);

        Record r2 = excelFile.next();
        Assert.verify(r2.getType() == Type.MSODRAWING ||
                      r2.getType() == Type.CONTINUE);
        if (r2.getType() == Type.MSODRAWING || r2.getType() == Type.CONTINUE)
        {
          MsoDrawingRecord mso = new MsoDrawingRecord(r2);
          checkBox.addMso(mso);
          r2 = excelFile.next();
        }

        Assert.verify(r2.getType() == Type.TXO);
        TextObjectRecord txo = new TextObjectRecord(r2);
        checkBox.setTextObject(txo);

        if (txo.getTextLength() == 0)
        {
          return;
        }

        r2 = excelFile.next();
        Assert.verify(r2.getType() == Type.CONTINUE);
        ContinueRecord text = new ContinueRecord(r2);
        checkBox.setText(text);

        r2 = excelFile.next();
        if (r2.getType() == Type.CONTINUE)
        {
          ContinueRecord formatting = new ContinueRecord(r2);
          checkBox.setFormatting(formatting);
        }

        drawings.add(checkBox);

        return;
      }

      // Handle form buttons
      if (objRecord.getType() == ObjRecord.BUTTON)
      {
        if (drawingData == null)
        {
          drawingData = new DrawingData();
        }

        Button button = new Button(msoRecord,
                                   objRecord,
                                   drawingData,
                                   workbook.getDrawingGroup(),
                                   workbookSettings);

        Record r2 = excelFile.next();
        Assert.verify(r2.getType() == Type.MSODRAWING ||
                      r2.getType() == Type.CONTINUE);
        if (r2.getType() == Type.MSODRAWING ||
            r2.getType() == Type.CONTINUE)
        {
          MsoDrawingRecord mso = new MsoDrawingRecord(r2);
          button.addMso(mso);
          r2 = excelFile.next();
        }

        Assert.verify(r2.getType() == Type.TXO);
        TextObjectRecord txo = new TextObjectRecord(r2);
        button.setTextObject(txo);

        r2 = excelFile.next();
        Assert.verify(r2.getType() == Type.CONTINUE);
        ContinueRecord text = new ContinueRecord(r2);
        button.setText(text);

        r2 = excelFile.next();
        if (r2.getType() == Type.CONTINUE)
        {
          ContinueRecord formatting = new ContinueRecord(r2);
          button.setFormatting(formatting);
        }

        drawings.add(button);

        return;
      }

      // Non-supported types which have multiple record types
      if (objRecord.getType() == ObjRecord.TEXT)
      {
        LOGGER.warn(objRecord.getType() + " Object on sheet \"" +
                    sheet.getName() +
                    "\" not supported - omitting");

        // Still need to add the drawing data to preserve the hierarchy
        if (drawingData == null)
        {
          drawingData = new DrawingData();

        }
        drawingData.addData(msoRecord.getData());

        Record r2 = excelFile.next();
        Assert.verify(r2.getType() == Type.MSODRAWING ||
                      r2.getType() == Type.CONTINUE);
        if (r2.getType() == Type.MSODRAWING ||
            r2.getType() == Type.CONTINUE)
        {
          MsoDrawingRecord mso = new MsoDrawingRecord(r2);
          drawingData.addRawData(mso.getData());
          r2 = excelFile.next();
        }

        Assert.verify(r2.getType() == Type.TXO);

        if (workbook.getDrawingGroup() != null) // can be null for Excel 95
        {
          workbook.getDrawingGroup().setDrawingsOmitted(msoRecord,
                                                        objRecord);
        }

        return;
      }

      // Handle other types
      if (objRecord.getType() != ObjRecord.CHART)
      {
        LOGGER.warn(objRecord.getType() + " Object on sheet \"" +
                    sheet.getName() +
                    "\" not supported - omitting");

        // Still need to add the drawing data to preserve the hierarchy
        if (drawingData == null)
        {
          drawingData = new DrawingData();
        }

        drawingData.addData(msoRecord.getData());

        if (workbook.getDrawingGroup() != null) // can be null for Excel 95
        {
          workbook.getDrawingGroup().setDrawingsOmitted(msoRecord,
                                                        objRecord);
        }

      }
    }
    catch (DrawingDataException e)
    {
      LOGGER.warn(e.getMessage() +
                  "...disabling drawings for the remainder of the workbook");
      workbookSettings.setDrawingsDisabled(true);
    }
  }

  /**
   * Gets the drawing data - for use as part of the Escher debugging tool
   */
  DrawingData getDrawingData()
  {
    return drawingData;
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
