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
    int chars = IntegerHelper.getInt(d[0], d[1]);

    if (chars == 0)
      return "";

    int optionFlags = d[2];

    int pos = 3;
    if ((optionFlags & 0xf) != optionFlags)
    {
      // Uh oh - looks like a plain old string, not unicode
      // Recalculate all the positions
      chars = IntegerHelper.getInt(d[0], (byte) 0);
      optionFlags = d[1];
      pos = 2;
    }

    // See if it is an extended string
    boolean extendedString = ((optionFlags & 0x04) != 0);

    // See if string contains formatting information
    boolean richString = ((optionFlags & 0x08) != 0);

    if (richString)
      pos += 2;

    if (extendedString)
      pos += 4;

    return StringHelper.readBiff8String(d);
  }
}
