package jxl.write.biff;

import java.io.IOException;
import java.nio.file.Files;
import jxl.WorkbookSettings;
import jxl.biff.FormattingRecords;
import static jxl.biff.Type.HORIZONTALPAGEBREAKS;
import jxl.write.*;
import static org.junit.Assert.*;
import org.junit.Test;

/**
 * created 2015-03-08
 * @author Jan Schlößin
 */
public class HorizontalPageBreaksRecordTest {
  
  @Test
  public void testCreationOfBiff8() {
    int [] pageBreaks = new int[] {13, 18};
    HorizontalPageBreaksRecord wpb = new HorizontalPageBreaksRecord(pageBreaks);
    assertArrayEquals(new byte[]{(byte) HORIZONTALPAGEBREAKS.value, 0, 14, 0, 2, 0, 13, 0, 0, 0, (byte) 255, (byte) 255, 18, 0, 0, 0, (byte) 255, (byte) 255},
            wpb.getBytes());
  }
  
  @Test
  public void testInsertionOfRowBreaks() {
    WritableSheetImpl w = new WritableSheetImpl("a", null, null, null, null, null);
    assertArrayEquals(new int[] {}, w.getRowPageBreaks());
    w.addRowPageBreak(5);
    assertArrayEquals(new int[] {5}, w.getRowPageBreaks());
  }
  
  @Test
  public void testDoubleInsertionOfRowBreaks() {
    WritableSheetImpl w = new WritableSheetImpl("a", null, null, null, null, null);
    w.addRowPageBreak(5);
    w.addRowPageBreak(5);
    assertArrayEquals(new int[] {5}, w.getRowPageBreaks());
  }
  
  @Test
  public void testInsertionOfRow() throws WriteException, IOException {
    WritableSheetImpl w = new WritableSheetImpl("[a]", null, new FormattingRecords(null), new SharedStrings(), new WorkbookSettings(), new WritableWorkbookImpl(Files.newOutputStream(Files.createTempFile(null, null)), true, new WorkbookSettings()));
    w.addCell(new Blank(10, 10));
    w.addRowPageBreak(5);
    w.addRowPageBreak(6);
    w.insertRow(6);
    assertArrayEquals(new int[] {5,7}, w.getRowPageBreaks());
    w.insertRow(0);
    assertArrayEquals(new int[] {6,8}, w.getRowPageBreaks());
  }
  
  @Test
  public void testRemovalOfRow() throws WriteException, IOException {
    WritableSheetImpl w = new WritableSheetImpl("[a]", null, new FormattingRecords(null), new SharedStrings(), new WorkbookSettings(), new WritableWorkbookImpl(Files.newOutputStream(Files.createTempFile(null, null)), true, new WorkbookSettings()));
    w.addCell(new Blank(10, 10));
    w.addRowPageBreak(6);
    w.addRowPageBreak(8);
    w.removeRow(0);
    assertArrayEquals(new int[] {5,7}, w.getRowPageBreaks());
    w.removeRow(6);
    assertArrayEquals(new int[] {5,6}, w.getRowPageBreaks());
  }
  
}
