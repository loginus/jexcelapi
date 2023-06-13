package jxl.read.biff;

import jxl.biff.*;
import jxl.biff.formula.*;

/**
 * created 29.05.2020
 * @author jan
 */
public class StringFormulaRecordBiff8 extends StringFormulaRecord {

  public StringFormulaRecordBiff8(Record t, File excelFile,
          FormattingRecords fr, ExternalSheet es, WorkbookMethods nt,
          SheetImpl si) {
    super(t, excelFile, fr, es, nt, si);
    value = readString(stringData);
  }

  public StringFormulaRecordBiff8(Record t, FormattingRecords fr,
          ExternalSheet es, WorkbookMethods nt, SheetImpl si) {
    super(t, fr, es, nt, si);
  }

  /**
   * Reads in the string
   *
   * @param d the data
   * @param ws the workbook settings
   */
  private String readString(byte[] d)
  {
    return StringHelper.readBiff8String(d);
  }
}
