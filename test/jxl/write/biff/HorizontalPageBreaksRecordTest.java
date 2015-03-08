package jxl.write.biff;

import static jxl.biff.Type.HORIZONTALPAGEBREAKS;
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
  
}
