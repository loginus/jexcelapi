package jxl.biff;

import static org.hamcrest.core.Is.is;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author jan
 */
public class StringHelperTest {

  @Test
  public void testReadingOfUtf16() {
    byte[] oe_ABreve = new byte[] {(byte)2, (byte)0, (byte)1, (byte)0xf6, (byte)0, (byte)2, (byte)1};

    assertThat(StringHelper.readBiff8String(oe_ABreve), is("öĂ"));
  }

  @Test
  public void testReadingOfCompressedUft16() {
    byte[] oe_ae_ue_sz = new byte[] {(byte)4, (byte)0, (byte)0, (byte)0xf6, (byte)0xe4, (byte)0xfc, (byte)0xdf};

    assertThat(StringHelper.readBiff8String(oe_ae_ue_sz), is("öäüß"));
  }

}
