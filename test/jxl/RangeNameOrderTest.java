
package jxl;

import java.io.IOException;
import java.net.URI;
import java.nio.file.*;
import java.util.List;
import java.util.logging.*;
import static java.util.stream.Collectors.toList;
import java.util.stream.Stream;
import static jxl.Workbook.getWorkbook;
import jxl.biff.*;
import jxl.biff.formula.FormulaException;
import jxl.read.biff.*;
import jxl.write.*;
import static org.junit.Assert.*;
import org.junit.Test;

/**
 * created 08.07.2021
 * @author jan
 */
public class RangeNameOrderTest {

  @Test
  public void afterWriting_RecordNameOrderIsSame() throws IOException, BiffException {
    Path origRangeNameWithPrintArea = Path.of(URI.create(getClass().getResource("/testdata/printAreaAndNamedRanges.xls").toString()));
    Path tempFile = Files.createTempFile("test", ".xls");
    List<String> origRecordNames;
    try (Workbook wb = getWorkbook(origRangeNameWithPrintArea);
         WritableWorkbook ww = Workbook.createWorkbook(tempFile, wb)) {
      origRecordNames = recordNames((WorkbookParser) wb);
      ww.write();
    }

    List<String> writtenRecordNames;
    try (Workbook wb = getWorkbook(tempFile)) {
      writtenRecordNames = recordNames((WorkbookParser) wb);
    }

    // the ranges are referenced in formulas as index in the name table
    // so the order have to be stable within reading and writing
    // even in respect of print area.
    assertEquals(origRecordNames, writtenRecordNames);
    Files.delete(tempFile);
  }

  private List<String> recordNames(WorkbookParser wb) {
    return wb.getNameRecords().stream()
            .map(record -> record.getName() == null ? record.getBuiltInName().getName() : record.getName())
            .collect(toList());
  }

}
