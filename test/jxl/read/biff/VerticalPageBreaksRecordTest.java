package jxl.read.biff;

import static jxl.biff.Type.VERTICALPAGEBREAKS;
import static jxl.read.biff.VerticalPageBreaksRecord.biff7;
import static org.junit.Assert.*;
import org.junit.Test;

/**
 * created 2015-03-08
 * @author Jan Schlößin
 */
public class VerticalPageBreaksRecordTest {
  
  @Test
  public void testCreationFromBiff7() {
    Record r = new TestRecord(
            new byte[]{(byte) VERTICALPAGEBREAKS.value, 0, 14, 0},
            new byte[]{2, 0, 2, 0, 8, 0});
    VerticalPageBreaksRecord pb = new VerticalPageBreaksRecord(r, biff7);
    assertEquals(2, pb.getColumnBreaks().length);
    assertEquals(2, pb.getColumnBreaks()[0]);
    assertEquals(8, pb.getColumnBreaks()[1]);
  }
  
  @Test
  public void testCreationFromBiff8() {
    Record r = new TestRecord(
            new byte[]{(byte) VERTICALPAGEBREAKS.value, 0, 14, 0},
            new byte[]{2, 0, 2, 0, 0, 0, (byte) 255, (byte) 255, 8, 0, 0, 0, (byte) 255, (byte) 255});
    VerticalPageBreaksRecord pb = new VerticalPageBreaksRecord(r);
    assertEquals(2, pb.getColumnBreaks().length);
    assertEquals(2, pb.getColumnBreaks()[0]);
    assertEquals(8, pb.getColumnBreaks()[1]);
  }
  
}
