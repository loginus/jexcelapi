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
  
  @Test
  public void testCreationFromBiff7() {
    Record r = new TestRecord(
            new byte[]{(byte) HORIZONTALPAGEBREAKS.value, 0, 14, 0},
            new byte[]{2, 0, 13, 0, 18, 0});
    HorizontalPageBreaksRecord pb = new HorizontalPageBreaksRecord(r, biff7);
    assertEquals(2, pb.getRowBreaks().length);
    assertEquals(13, pb.getRowBreaks()[0]);
    assertEquals(18, pb.getRowBreaks()[1]);
  }
  
  @Test
  public void testCreationFromBiff8() {
    Record r = new TestRecord(
            new byte[]{(byte) HORIZONTALPAGEBREAKS.value, 0, 14, 0},
            new byte[]{2, 0, 13, 0, 0, 0, (byte) 255, (byte) 255, 18, 0, 0, 0, (byte) 255, (byte) 255});
    HorizontalPageBreaksRecord pb = new HorizontalPageBreaksRecord(r);
    assertEquals(2, pb.getRowBreaks().length);
    assertEquals(13, pb.getRowBreaks()[0]);
    assertEquals(18, pb.getRowBreaks()[1]);
  }
  
}
