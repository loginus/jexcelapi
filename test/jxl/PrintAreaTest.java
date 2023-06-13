package jxl;

import java.io.*;
import java.net.*;
import java.nio.file.*;
import java.util.Arrays;
import java.util.stream.Stream;
import static jxl.Workbook.getWorkbook;
import jxl.read.biff.*;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import static org.hamcrest.core.Is.is;
import static org.junit.Assert.assertThat;
import org.junit.Test;

/**
 * created 2021-01-30
 * @author jan
 */
public class PrintAreaTest {

  @Test
  public void insertRow() throws IOException, BiffException, RowsExceededException {
    Path source = Path.of(URI.create(getClass().getResource("/testdata/printArea.xls").toString()));
    Path destination = Files.createTempFile("test", ".xls");
    try (
            Workbook wb = getWorkbook(source);
            WritableWorkbook ww = Workbook.createWorkbook(destination, wb)
            ) {
      WritableSheet sheet = ww.getSheet(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(1));
      sheet.insertRow(2);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(1));
      sheet.insertRow(1);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(2));
      sheet.insertRow(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(1));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(3));

      ww.write();
    }
    Files.delete(destination);
  }

  @Test
  public void removeRow() throws IOException, BiffException, RowsExceededException {
    Path source = Path.of(URI.create(getClass().getResource("/testdata/printArea.xls").toString()));
    Path destination = Files.createTempFile("test", ".xls");
    try (
            Workbook wb = getWorkbook(source);
            WritableWorkbook ww = Workbook.createWorkbook(destination, wb)
            ) {
      WritableSheet sheet = ww.getSheet(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(1));
      sheet.removeRow(2);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(1));
      sheet.removeRow(1);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(0));
      sheet.removeRow(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getRow(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getRow(), is(0));

      ww.write();
    }
    Files.delete(destination);
  }

  @Test
  public void insertColumn() throws IOException, BiffException, RowsExceededException {
    Path source = Path.of(URI.create(getClass().getResource("/testdata/printArea.xls").toString()));
    Path destination = Files.createTempFile("test", ".xls");
    try (
            Workbook wb = getWorkbook(source);
            WritableWorkbook ww = Workbook.createWorkbook(destination, wb)
            ) {
      WritableSheet sheet = ww.getSheet(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(1));
      sheet.insertColumn(2);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(1));
      sheet.insertColumn(1);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(2));
      sheet.insertColumn(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(1));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(3));

      ww.write();
    }
    Files.delete(destination);
  }

  @Test
  public void removeColumn() throws IOException, BiffException, RowsExceededException {
    Path source = Path.of(URI.create(getClass().getResource("/testdata/printArea.xls").toString()));
    Path destination = Files.createTempFile("test", ".xls");
    try (
            Workbook wb = getWorkbook(source);
            WritableWorkbook ww = Workbook.createWorkbook(destination, wb)
            ) {
      WritableSheet sheet = ww.getSheet(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(1));
      sheet.removeColumn(2);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(1));
      sheet.removeColumn(1);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(0));
      sheet.removeColumn(0);
      assertThat(sheet.getSettings().getPrintArea().getTopLeft().getColumn(), is(0));
      assertThat(sheet.getSettings().getPrintArea().getBottomRight().getColumn(), is(0));

      ww.write();
    }
    Files.delete(destination);
  }

}
