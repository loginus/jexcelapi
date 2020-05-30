package jxl.biff;

/**
 * created 2020-05-30
 * @author jan
 */
public enum Version {
  BIFF2 { @Override String getExcelVersion() { return "Excel 2.x"; } },
  BIFF3 { @Override String getExcelVersion() { return "Excel 3.0"; } },
  BIFF4 { @Override String getExcelVersion() { return "Excel 4.0"; } },
  BIFF5 { @Override String getExcelVersion() { return "Excel 5.0"; } },
  BIFF7 { @Override String getExcelVersion() { return "Excel 95/7.0"; } },
  BIFF8 { @Override String getExcelVersion() { return "Excel 97-2003/8-11"; } };

  abstract String getExcelVersion();
}
