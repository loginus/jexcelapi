package jxl.write.biff;

import java.io.IOException;
import java.nio.file.Files;
import jxl.WorkbookSettings;
import jxl.biff.FormattingRecords;
import static jxl.biff.Type.VERTICALPAGEBREAKS;
import jxl.write.*;
import static org.junit.Assert.*;
import org.junit.Test;

/**
 * created 2015-03-08
 * @author Jan Schlößin
 */
public class VerticalPageBreaksRecordTest {
  
  @Test
  public void testCreationOfBiff8() {
    int [] pageBreaks = new int[] {2, 8};
    VerticalPageBreaksRecord wpb = new VerticalPageBreaksRecord(pageBreaks);
    assertArrayEquals(new byte[]{(byte) VERTICALPAGEBREAKS.value, 0, 14, 0, 2, 0, 2, 0, 0, 0, (byte) 255, (byte) 255, 8, 0, 0, 0, (byte) 255, (byte) 255},
            wpb.getBytes());
  }
  
  @Test
  public void testInsertionOfColumnBreaks() {
    WritableSheetImpl w = new WritableSheetImpl("a", null, null, null, null, null);
    assertArrayEquals(new int[] {}, w.getColumnPageBreaks());
    w.addColumnPageBreak(5);
    assertArrayEquals(new int[] {5}, w.getColumnPageBreaks());
  }
  
  @Test
  public void testDoubleInsertionOfColumnBreaks() {
    WritableSheetImpl w = new WritableSheetImpl("a", null, null, null, null, null);
    w.addColumnPageBreak(5);
    w.addColumnPageBreak(5);
    assertArrayEquals(new int[] {5}, w.getColumnPageBreaks());
  }
  
  @Test
  public void testInsertionOfColumn() throws WriteException, IOException {
    WritableSheetImpl w = new WritableSheetImpl("[a]", null, new FormattingRecords(null), new SharedStrings(), new WorkbookSettings(), new WritableWorkbookImpl(Files.newOutputStream(Files.createTempFile(null, null)), true, new WorkbookSettings()));
    w.addCell(new Blank(10, 10));
    w.addColumnPageBreak(5);
    w.addColumnPageBreak(6);
    w.insertColumn(6);
    assertArrayEquals(new int[] {5,7}, w.getColumnPageBreaks());
    w.insertColumn(0);
    assertArrayEquals(new int[] {6,8}, w.getColumnPageBreaks());
  }
  
  @Test
  public void testRemovalOfColumn() throws WriteException, IOException {
    WritableSheetImpl w = new WritableSheetImpl("[a]", null, new FormattingRecords(null), new SharedStrings(), new WorkbookSettings(), new WritableWorkbookImpl(Files.newOutputStream(Files.createTempFile(null, null)), true, new WorkbookSettings()));
    w.addCell(new Blank(10, 10));
    w.addColumnPageBreak(6);
    w.addColumnPageBreak(8);
    w.removeColumn(0);
    assertArrayEquals(new int[] {5,7}, w.getColumnPageBreaks());
    w.removeColumn(6);
    assertArrayEquals(new int[] {5,6}, w.getColumnPageBreaks());
  }
  
}
