package jxl.biff;

import java.io.*;
import java.net.*;
import java.nio.file.*;
import jxl.*;
import static jxl.Workbook.getWorkbook;
import jxl.biff.formula.*;
import jxl.read.biff.*;
import jxl.write.*;
import static org.hamcrest.core.Is.is;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * created 2020-05-22
 * @author jan
 */
public class StringHelperTest {

  @Test
  public void testReadingOfBiff8String() {
    int[] öĂ = new int[] {
      2, 0,    // length
      1,       // uncompressed
      0xf6, 0, // ö
      2, 1     // Ă
    };

    assertThat(StringHelper.readBiff8String(toBytes(öĂ)), is("öĂ"));
  }

  @Test
  public void testReadingOfCompressedBiff8String() {
    int[] öäüß = new int[] {
      4, 0, // length
      0,    // compressed
      0xf6, // ö
      0xe4, // ä
      0xfc, // ü
      0xdf  // ß
    };

    assertThat(StringHelper.readBiff8String(toBytes(öäüß)), is("öäüß"));
  }

  @Test
  public void testReadingOfShortBiff8String() {
    int[] öĂ = new int[] {
      1,       // uncompressed
      0xf6, 0, // ö
      2, 1     // Ă
    };

    assertThat(StringHelper.readShortBiff8String(toBytes(öĂ)), is("öĂ"));
  }

  @Test
  public void testReadingOfShortCompressedBiff8String() {
    int[] öäüß = new int[] {
      0,    // compressed
      0xf6, // ö
      0xe4, // ä
      0xfc, // ü
      0xdf  // ß
    };

    assertThat(StringHelper.readShortBiff8String(toBytes(öäüß)), is("öäüß"));
  }

  private byte[] toBytes(int[] ints) {
    byte[] bytes = new byte[ints.length];
    for (var i = 0; i < ints.length; i++)
      bytes[i] = (byte) ints[i];

    return bytes;
  }

  @Test
  public void testUtf16InComments() throws IOException, BiffException {
    Path source = Path.of(URI.create(getClass().getResource("/testdata/utf16InComments.xls").toString()));
    Path destination = Files.createTempFile("test", ".xls");
    try (
            Workbook wb = getWorkbook(source);
            WritableWorkbook ww = Workbook.createWorkbook(destination, wb)
            ) {
      Sheet s = wb.getSheet(0);
      assertThat(s.getCell(0, 0).getCellFeatures().getComment(), is("äöüß"));
      assertThat(s.getCell(0, 1).getCellFeatures().getComment(), is("Ă"));

      ww.write();
      WritableSheet ws = ww.getSheet(0);
      assertThat(ws.getCell(0, 0).getCellFeatures().getComment(), is("äöüß"));
      assertThat(ws.getCell(0, 1).getCellFeatures().getComment(), is("Ă"));
    }
    Files.delete(destination);
  }

  @Test
  public void testUft16InFormulas() throws IOException, BiffException, FormulaException {
    Path source = Path.of(URI.create(getClass().getResource("/testdata/utf16InFormulas.xls").toString()));
    Path destination = Files.createTempFile("test", ".xls");
    try (
            Workbook wb = getWorkbook(source);
            WritableWorkbook ww = Workbook.createWorkbook(destination, wb)
            ) {
      Sheet s = wb.getSheet(0);
      assertThat(((StringFormulaCell) s.getCell(0, 0)).getString(), is("öäüß"));
      assertThat(((StringFormulaCell) s.getCell(0, 1)).getString(), is("Ă"));

      ww.write();
      WritableSheet ws = ww.getSheet(0);
      assertThat(((StringFormulaCell) ws.getCell(0, 0)).getString(), is("öäüß"));
      assertThat(((StringFormulaCell) ws.getCell(0, 1)).getString(), is("Ă"));
    }
  }

}
