package jxl.write.biff;

import java.io.IOException;
import java.nio.file.*;
import jxl.*;
import jxl.biff.FormattingRecords;
import static jxl.biff.Type.VERTICALPAGEBREAKS;
import jxl.read.biff.BiffException;
import static jxl.read.biff.VerticalPageBreaksRecordTest.biff8pb;
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
    VerticalPageBreaksRecord wpb = new VerticalPageBreaksRecord();
    wpb.setColumnBreaks(biff8pb);
    assertArrayEquals(new byte[]{(byte) VERTICALPAGEBREAKS.value, 0, 14, 0, 2, 0, 2, 0, 0, 0, (byte) 255, (byte) 255, 8, 0, 0, 0, (byte) 255, (byte) 255},
            wpb.getBytes());
  }
  
  @Test
  public void testInsertionOfColumnBreaks() {
    WritableSheetImpl w = new WritableSheetImpl("a", null, null, null, null, null);
    assertTrue(w.getColumnPageBreaks().getColumnBreaks().isEmpty());
    w.addColumnPageBreak(5);
    assertEquals(5, (int) w.getColumnPageBreaks().getColumnBreaks().get(0));
  }
  
  @Test
  public void testDoubleInsertionOfColumnBreaks() {
    WritableSheetImpl w = new WritableSheetImpl("a", null, null, null, null, null);
    w.addColumnPageBreak(5);
    assertEquals(1, w.getColumnPageBreaks().getColumnBreaks().size());
    w.addColumnPageBreak(5);
    assertEquals(1, w.getColumnPageBreaks().getColumnBreaks().size());
  }
  
  @Test
  public void testInsertionOfColumn() throws WriteException, IOException {
    WritableSheetImpl w = new WritableSheetImpl("[a]", null, new FormattingRecords(null), new SharedStrings(), new WorkbookSettings(), new WritableWorkbookImpl(Files.newOutputStream(Files.createTempFile(null, null)), true, new WorkbookSettings()));
    w.addCell(new Blank(10, 10));
    w.addColumnPageBreak(5);
    w.addColumnPageBreak(6);
    w.insertColumn(6);
    assertArrayEquals(new Integer[] {5,7}, w.getColumnPageBreaks().getColumnBreaks().toArray());
    w.insertColumn(0);
    assertArrayEquals(new Integer[] {6,8}, w.getColumnPageBreaks().getColumnBreaks().toArray());
  }
  
  @Test
  public void testRemovalOfColumn() throws WriteException, IOException {
    WritableSheetImpl w = new WritableSheetImpl("[a]", null, new FormattingRecords(null), new SharedStrings(), new WorkbookSettings(), new WritableWorkbookImpl(Files.newOutputStream(Files.createTempFile(null, null)), true, new WorkbookSettings()));
    w.addCell(new Blank(10, 10));
    w.addColumnPageBreak(6);
    w.addColumnPageBreak(8);
    w.removeColumn(0);
    assertArrayEquals(new Integer[] {5,7}, w.getColumnPageBreaks().getColumnBreaks().toArray());
    w.removeColumn(6);
    assertArrayEquals(new Integer[] {5,6}, w.getColumnPageBreaks().getColumnBreaks().toArray());
    w.removeColumn(6);
    assertArrayEquals(new Integer[] {5}, w.getColumnPageBreaks().getColumnBreaks().toArray());
  }
  
  @Test
  public void testIntegration_CopyOfWorkbookWithColumnBreak() throws WriteException, IOException, BiffException {
    Path tempSource = Files.createTempFile(null, ".xls");
    try (WritableWorkbook wwb = Workbook.createWorkbook(tempSource)) {
      WritableSheet sheet = wwb.createSheet("test", 0);
      sheet.addColumnPageBreak(6);
      wwb.write();
    }
    Path tempDest = Files.createTempFile(null, ".xls");
    try (Workbook wb = Workbook.getWorkbook(tempSource);
            WritableWorkbook wwb = Workbook.createWorkbook(tempDest, wb)) {
      wwb.write();
    }
    
    Files.deleteIfExists(tempSource);
    Files.deleteIfExists(tempDest);
  }
  
  @Test
  public void testIntegration_CopyOfWorkbookWithoutColumnBreak() throws WriteException, IOException, BiffException {
    Path tempSource = Files.createTempFile(null, ".xls");
    try (WritableWorkbook wwb = Workbook.createWorkbook(tempSource)) {
      wwb.createSheet("test", 0);
      wwb.write();
    }
    Path tempDest = Files.createTempFile(null, ".xls");
    try (Workbook wb = Workbook.getWorkbook(tempSource);
            WritableWorkbook wwb = Workbook.createWorkbook(tempDest, wb)) {
      wwb.write();
    }
    
    Files.deleteIfExists(tempSource);
    Files.deleteIfExists(tempDest);
  }
  
}
