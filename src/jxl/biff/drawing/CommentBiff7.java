
package jxl.biff.drawing;

import jxl.*;
import jxl.biff.*;
import jxl.common.*;

/**
 * created 29.05.2020
 * @author jan
 */
public class CommentBiff7 extends CommentBiff8 {

  /**
   * The workbook settings
   */
  private final WorkbookSettings workbookSettings;

  public CommentBiff7(
          MsoDrawingRecord msorec,
          ObjRecord obj,
          DrawingData dd,
          DrawingGroup dg,
          WorkbookSettings workbookSettings) {
    super(msorec, obj, dd, dg);
    this.workbookSettings = workbookSettings;
  }

  public CommentBiff7(
          DrawingGroupObject dgo,
          DrawingGroup dg,
          WorkbookSettings workbookSettings) {
    super(dgo, dg);
    this.workbookSettings = workbookSettings;
  }

  public CommentBiff7(
          String txt, int c, int r, WorkbookSettings workbookSettings) {
    super(txt, c, r);
    this.workbookSettings = workbookSettings;
  }

  @Override
  public String getText() {
    if (commentText == null)
    {
      Assert.verify(text != null);

      byte[] td = text.getData();
      if (td[0] == 0)
      {
        commentText = StringHelper.getString
          (td, td.length - 1, 1, workbookSettings);
    }
      else
      {
        commentText = StringHelper.getUnicodeString
          (td, 1, (td.length - 1) / 2);
      }
    }

    return commentText;
  }

}
