package jxl.write.biff;

import static jxl.biff.Type.VERTICALPAGEBREAKS;
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
  
}
