package jxl.read.biff;

import static jxl.biff.Type.HORIZONTALPAGEBREAKS;
import static jxl.read.biff.HorizontalPageBreaksRecord.biff7;
import static org.junit.Assert.*;
import org.junit.Test;

/**
 * created 2015-03-08
 * @author Jan Schlößin
 */
public class HorizontalPageBreaksRecordTest {
  
  public static final Record biff8HorizontalPageBreakRecord = new TestRecord(
          new byte[]{(byte) HORIZONTALPAGEBREAKS.value, 0, 14, 0},
          new byte[]{2, 0, 13, 0, 0, 0, (byte) 255, (byte) 255, 18, 0, 0, 0, (byte) 255, (byte) 255});
  public static final HorizontalPageBreaksRecord biff8pb = new HorizontalPageBreaksRecord(biff8HorizontalPageBreakRecord);
  
  @Test
  public void testCreationFromBiff7() {
    Record r = new TestRecord(
            new byte[]{(byte) HORIZONTALPAGEBREAKS.value, 0, 14, 0},
            new byte[]{2, 0, 13, 0, 18, 0});
    HorizontalPageBreaksRecord pb = new HorizontalPageBreaksRecord(r, biff7);
    assertEquals(2, pb.getRowBreaks().size());
    assertEquals(13, (int) pb.getRowBreaks().get(0));
    assertEquals(18, (int) pb.getRowBreaks().get(1));
  }
  
  @Test
  public void testCreationFromBiff8() {
    assertEquals(2, biff8pb.getRowBreaks().size());
    assertEquals(13, (int) biff8pb.getRowBreaks().get(0));
    assertEquals(18, (int) biff8pb.getRowBreaks().get(1));
  }
  
}
