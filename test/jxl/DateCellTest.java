package jxl;

import java.io.IOException;
import java.nio.file.*;
import java.time.*;
import java.util.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import static jxl.write.DateTime.GMT;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * created 2015-08-26
 * @author jan
 */
public class DateCellTest {
  
  @Test
  public void testDateWritingAndRewriting() throws IOException, WriteException, BiffException {
    Date summer = new GregorianCalendar(2015, 8, 26, 17, 15, 39).getTime();
    Date winter = new GregorianCalendar(2015, 1, 26, 17, 15, 39).getTime();
    Path file = Files.createTempFile("test", ".xls");
    try (WritableWorkbook wb = Workbook.createWorkbook(file)) {
      WritableSheet sheet = wb.createSheet("test", 0);
      sheet.addCell(new DateTime(0, 0, summer, new WritableCellFormat(DateFormats.FORMAT9)));
      sheet.addCell(new DateTime(0, 1, winter, new WritableCellFormat(DateFormats.FORMAT9)));
      wb.write();
      assertEquals(CellType.DATE, sheet.getCell(0, 0).getType());
      assertEquals(summer, ((DateCell) sheet.getCell(0, 0)).getDate());
      assertEquals(summer.toString(), ((DateCell) sheet.getCell(0, 0)).getContents());
      assertEquals(winter, ((DateCell) sheet.getCell(0, 1)).getDate());
      assertEquals(winter.toString(), ((DateCell) sheet.getCell(0, 1)).getContents());
    }
    try (Workbook wb = Workbook.getWorkbook(file)) {
      Sheet sheet = wb.getSheet(0);
      assertEquals(CellType.DATE, sheet.getCell(0, 0).getType());
      assertEquals(summer, ((DateCell) sheet.getCell(0, 0)).getDate());
      assertEquals("9/26/15 17:15", ((DateCell) sheet.getCell(0, 0)).getContents());
      assertEquals(winter, ((DateCell) sheet.getCell(0, 1)).getDate());
      assertEquals("2/26/15 17:15", ((DateCell) sheet.getCell(0, 1)).getContents());
    }
    
  }
  
}
