package jxl.write.biff;

import java.io.IOException;
import java.nio.file.*;
import jxl.*;
import jxl.biff.EmptyCell;
import jxl.write.*;
import static org.hamcrest.core.Is.is;
import org.junit.*;
import static org.junit.Assert.*;

/**
 * created 2015-08-20
 * @author jan
 */
public class WritableSheetImplTest {

  private Path tempFile;
  private WritableWorkbook workbook;
  private WritableSheet sheet;

  @Before
  public void setUp() throws IOException {
    tempFile = Files.createTempFile(null, ".xls");
    workbook = Workbook.createWorkbook(tempFile);
    sheet = workbook.createSheet("a", 0);
  }

  @After
  public void tearDown() throws IOException {
    sheet = null;
    workbook.write();
    workbook.close();
    workbook = null;
    Files.delete(tempFile);
    tempFile = null;
  }

  @Test
  public void testAddingAnEmptyCell() throws WriteException, IOException {
    sheet.addCell(new Label(0, 0, "test"));
    assertEquals("test", sheet.getCell(0, 0).getContents());
    sheet.addCell(new EmptyCell(0, 0));
    assertEquals(CellType.EMPTY, sheet.getCell(0, 0).getType());
  }

  @Test
  public void addingALine_MovesFollowingCellsDown() throws IOException, WriteException {
    Label label = new Label(0, 0, "content");
    sheet.addCell(label);
    assertThat(label.getRow(), is(0));
    sheet.insertRow(0);
    assertThat(label.getRow(), is(1));
  }

  @Test
  public void addingALineInAnAlmostFullSheet_MovesFollowingCellsDown() throws IOException, WriteException {
    Label label = new Label(0, WritableSheetImpl.MAX_ROWS_PER_SHEET-2, "content");
    sheet.addCell(label);
    assertThat(label.getRow(), is(WritableSheetImpl.MAX_ROWS_PER_SHEET-2));
    sheet.insertRow(0);
    assertThat(label.getRow(), is(WritableSheetImpl.MAX_ROWS_PER_SHEET-1));
  }

  @Test
  public void addingALine_DoesntTouchTheCellsAbove() throws IOException, WriteException {
    Label label = new Label(0, 0, "content");
    sheet.addCell(label);
    assertThat(label.getRow(), is(0));
    sheet.insertRow(1);
    assertThat(label.getRow(), is(0));
  }

  @Test(expected = RowsExceededException.class)
  public void addingALineToAFullSheet_ThrowsException() throws IOException, WriteException {
    Label label = new Label(0, WritableSheetImpl.MAX_ROWS_PER_SHEET - 1, "content");
    sheet.addCell(label);
    sheet.insertRow(0);
  }

  @Test
  public void addingALine_MovesImagesDown() throws RowsExceededException {
    WritableImage image = new WritableImage(0, 0, 1, 1, new byte[] {});
    sheet.addImage(image);
    assertThat(image.getY(), is(0d));
    sheet.insertRow(0);
    assertThat(image.getY(), is(1d));
  }

  @Test
  public void addingALine_AndAnImageIsAlmostAtTheBottom_MovesImagesDown() throws RowsExceededException {
    WritableImage image = new WritableImage(0, WritableSheetImpl.MAX_ROWS_PER_SHEET - 3, 1, 1, new byte[] {});
    sheet.addImage(image);
    assertThat(image.getY(), is(WritableSheetImpl.MAX_ROWS_PER_SHEET - 3d));
    sheet.insertRow(0);
    assertThat(image.getY(), is(WritableSheetImpl.MAX_ROWS_PER_SHEET - 2d));
  }

  @Test
  public void addingALine_DoesntTouchTheImagesAbove() throws IOException, WriteException {
    WritableImage image = new WritableImage(0, 0, 1, 1, new byte[] {});
    sheet.addImage(image);
    assertThat(image.getY(), is(0d));
    sheet.insertRow(1);
    assertThat(image.getY(), is(0d));
  }

  @Test
  public void addingALine_AndAnImageIsAtTheSheetBottom_TheImageWillNotBeTouched() throws RowsExceededException, WriteException {
    WritableImage image = new WritableImage(0, WritableSheetImpl.MAX_ROWS_PER_SHEET - 2, 1, 1, new byte[] {});
    sheet.addImage(image);
    assertThat(image.getY(), is(WritableSheetImpl.MAX_ROWS_PER_SHEET - 2d));
    sheet.insertRow(0);
    assertThat(image.getY(), is(WritableSheetImpl.MAX_ROWS_PER_SHEET - 2d));
  }

}
