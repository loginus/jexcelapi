package jxl.biff;

import jxl.*;
import static jxl.biff.EncodedURLHelper.getEncodedURL;
import org.junit.Test;
import static org.junit.Assert.*;
import static org.hamcrest.core.Is.is;

/**
 * 2019-12-23
 * @author jan
 */
public class EncodedURLHelperTest {

  @Test
  public void testBiff4W_Biff8Encoding() {
    assertThat(getEncodedURL("[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01[ext.xls]Sheet1")));
    assertThat(getEncodedURL("sub\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01sub\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\02[ext.xls]Sheet1")));
    assertThat(getEncodedURL("\\sub\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\02sub\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("\\sub\\sub2\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\02sub\03sub2\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("..\\sub\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\04sub\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("D:\\sub\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\01Dsub\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("D:sub\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\01Dsub\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("\\\\pc\\sub\\[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\01@pc\03sub\03[ext.xls]Sheet1")));
    assertThat(getEncodedURL("http://www.example.org/[ext.xls]Sheet1", new WorkbookSettings()), is(convertFileName("\01\05\u0026http://www.example.org/[ext.xls]Sheet1")));
  }

  @Test
  public void testBiff2_Biff4Encoding() {
    assertThat(getEncodedURL("ext.xls", new WorkbookSettings()), is(convertFileName("\01ext.xls")));
    assertThat(getEncodedURL("sub\\ext.xls", new WorkbookSettings()), is(convertFileName("\01sub\03ext.xls")));
    assertThat(getEncodedURL("\\ext.xls", new WorkbookSettings()), is(convertFileName("\01\02ext.xls")));
    assertThat(getEncodedURL("\\sub\\ext.xls", new WorkbookSettings()), is(convertFileName("\01\02sub\03ext.xls")));
    assertThat(getEncodedURL("\\sub\\sub2\\ext.xls", new WorkbookSettings()), is(convertFileName("\01\02sub\03sub2\03ext.xls")));
    assertThat(getEncodedURL("..\\sub\\ext.xls", new WorkbookSettings()), is(convertFileName("\01\04sub\03ext.xls")));
    assertThat(getEncodedURL("D:\\sub\\ext.xls", new WorkbookSettings()), is(convertFileName("\01\01Dsub\03ext.xls")));
    assertThat(getEncodedURL("\\\\pc\\sub\\ext.xls", new WorkbookSettings()), is(convertFileName("\01\01@pc\03sub\03ext.xls")));
    assertThat(getEncodedURL("http://www.example.org/ext.xls", new WorkbookSettings()), is(convertFileName("\01\05\u001ehttp://www.example.org/ext.xls")));
  }

  private byte[] convertFileName(String filename) {
    return filename.substring(1).getBytes();
  }

}
