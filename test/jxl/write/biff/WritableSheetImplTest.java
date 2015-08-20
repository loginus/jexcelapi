package jxl.write.biff;

import java.io.IOException;
import java.nio.file.Files;
import jxl.*;
import jxl.biff.EmptyCell;
import jxl.write.*;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * created 2015-08-20
 * @author jan
 */
public class WritableSheetImplTest {
  
  @Test
  public void testAddingAnEmptyCell() throws WriteException, IOException {
    try (WritableWorkbook ww = Workbook.createWorkbook(Files.createTempFile(null, null))) {
      WritableSheet sheet = ww.createSheet("a", 0);
      sheet.addCell(new Label(0, 0, "test"));
      assertEquals("test", sheet.getCell(0, 0).getContents());
      sheet.addCell(new EmptyCell(0, 0));
      assertEquals(CellType.EMPTY, sheet.getCell(0, 0).getType());
    }
  }
  
}
