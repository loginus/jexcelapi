/*********************************************************************
*
*      Copyright (C) 2012 Jan Schloessin
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
*
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
* Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public
* License along with this library; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
***************************************************************************/

package jxl.write.biff;

import jxl.biff.*;

/**
 * Contains workbook level windowing attributes
 */
class Window1Record extends WritableRecordData
{
  private int horizontalPositionOfDocumentWindow;
  private int verticalPositionOfDocumentWindow;
  private int widthOfDocumentWindow;
  private int heightOfDocumentWindow;
  private boolean windowHidden;
  private boolean windowMinimized;
  private boolean horizontalScrollBarVisible;
  private boolean verticalScrollBarVisible;
  private boolean worksheetTabBarVisible;
  private int selectedSheet;
  private int firstVisibleTab;
  private int numberOfSelectedWorkSheets;
  private int widthOfWorksheetTabBar;
  
  /**
   * Constructor for compatibility with initial version from Andrew Khan
   */
  Window1Record(int selectedSheet)
  {
    this(360, 270, 14940, 9150, false, false, true, true, true,
            selectedSheet, 0, 1, 600);
  }

  /**
   * Assignment Constructor
   * 
   * @param horizontalPositionOfDocumentWindow Horizontal position of the document window (in twips = 1/20 of a point)
   * @param verticalPositionOfDocumentWindow Vertical position of the document window (in twips = 1/20 of a point)
   * @param widthOfDocumentWindow Width of the document window (in twips = 1/20 of a point)
   * @param heightOfDocumentWindow Height of the document window (in twips = 1/20 of a point)
   * @param windowHidden true when window is hidden, false when window is visible
   * @param windowMinimized true when window is minimized, false when window is open
   * @param horizontalScrollBarVisible true when horizontal scroll bar is visible, false when horizontal scroll bar hidden
   * @param verticalScrollBarVisible true when vertical scroll bar visible, false when vertical scroll bar hidden
   * @param worksheetTabBarVisible true when worksheet tab bar visible, false when worksheet tab bar hidden
   * @param selectedSheet Index to active (displayed) worksheet
   * @param firstVisibleTab Index of first visible tab in the worksheet tab bar
   * @param numberOfSelectedWorkSheets Number of selected worksheets (highlighted in the worksheet tab bar)
   * @param widthOfWorksheetTabBar Width of worksheet tab bar (in 1/1000 of window width). The remaining space is used by the horizontal scrollbar.
   */
  Window1Record(
          int horizontalPositionOfDocumentWindow,
          int verticalPositionOfDocumentWindow,
          int widthOfDocumentWindow,
          int heightOfDocumentWindow,
          boolean windowHidden,
          boolean windowMinimized,
          boolean horizontalScrollBarVisible,
          boolean verticalScrollBarVisible,
          boolean worksheetTabBarVisible,
          int selectedSheet,
          int firstVisibleTab,
          int numberOfSelectedWorkSheets,
          int widthOfWorksheetTabBar) {
    super(Type.WINDOW1);

    this.horizontalPositionOfDocumentWindow = horizontalPositionOfDocumentWindow;
    this.verticalPositionOfDocumentWindow = verticalPositionOfDocumentWindow;
    this.widthOfDocumentWindow = widthOfDocumentWindow;
    this.heightOfDocumentWindow = heightOfDocumentWindow;
    this.windowHidden = windowHidden;
    this.windowMinimized = windowMinimized;
    this.horizontalScrollBarVisible = horizontalScrollBarVisible;
    this.verticalScrollBarVisible = verticalScrollBarVisible;
    this.worksheetTabBarVisible = worksheetTabBarVisible;
    this.selectedSheet = selectedSheet;
    this.firstVisibleTab = firstVisibleTab;
    this.numberOfSelectedWorkSheets = numberOfSelectedWorkSheets;
    this.widthOfWorksheetTabBar = widthOfWorksheetTabBar;
  }

  /**
   * Gets the binary data for output to file
   * 
   * @return the binary data
   */
  @Override
  public byte[] getData()
  {
    byte[] data = new byte[18];
    IntegerHelper.getTwoBytes(horizontalPositionOfDocumentWindow, data, 0);
    IntegerHelper.getTwoBytes(verticalPositionOfDocumentWindow, data, 2);
    IntegerHelper.getTwoBytes(widthOfDocumentWindow, data, 4);
    IntegerHelper.getTwoBytes(heightOfDocumentWindow, data, 6);
		
    int optionFlag = 0;
    if (windowHidden)
			optionFlag |= 0x01;
		if (windowMinimized)
			optionFlag |= 0x02;
		if (horizontalScrollBarVisible)
			optionFlag |= 0x08;
		if (verticalScrollBarVisible)
			optionFlag |= 0x10;
		if (worksheetTabBarVisible)
			optionFlag |= 0x20;
    IntegerHelper.getTwoBytes(optionFlag, data, 8);
    
    IntegerHelper.getTwoBytes(selectedSheet, data, 10);
    IntegerHelper.getTwoBytes(firstVisibleTab, data, 12);
    IntegerHelper.getTwoBytes(numberOfSelectedWorkSheets, data, 14);
    IntegerHelper.getTwoBytes(widthOfWorksheetTabBar, data, 16);
    
    return data;
  }

  /**
   * @return the horizontal position of the document window (in twips = 1/20 of a point)
   */
  int getHorizontalPositionOfDocumentWindow() {
    return horizontalPositionOfDocumentWindow;
  }

  /**
   * @param horizontalPositionOfDocumentWindow Horizontal position of the document window (in twips = 1/20 of a point)
   */
  void setHorizontalPositionOfDocumentWindow(int horizontalPositionOfDocumentWindow) {
    this.horizontalPositionOfDocumentWindow = horizontalPositionOfDocumentWindow;
  }

  /**
   * @return the vertical position of the document window (in twips = 1/20 of a point)
   */
  int getVerticalPositionOfDocumentWindow() {
    return verticalPositionOfDocumentWindow;
  }

  /**
   * @param verticalPositionOfDocumentWindow Vertical position of the document window (in twips = 1/20 of a point)
   */
  void setVerticalPositionOfDocumentWindow(int verticalPositionOfDocumentWindow) {
    this.verticalPositionOfDocumentWindow = verticalPositionOfDocumentWindow;
  }

  /**
   * @return the width of the document window (in twips = 1/20 of a point)
   */
  int getWidthOfDocumentWindow() {
    return widthOfDocumentWindow;
  }

  /**
   * @param widthOfDocumentWindow Width of the document window (in twips = 1/20 of a point)
   */
  void setWidthOfDocumentWindow(int widthOfDocumentWindow) {
    this.widthOfDocumentWindow = widthOfDocumentWindow;
  }

  /**
   * @return the height of the document window (in twips = 1/20 of a point)
   */
  int getHeightOfDocumentWindow() {
    return heightOfDocumentWindow;
  }

  /**
   * @param heightOfDocumentWindow Height of the document window (in twips = 1/20 of a point)
   */
  void setHeightOfDocumentWindow(int heightOfDocumentWindow) {
    this.heightOfDocumentWindow = heightOfDocumentWindow;
  }

  /**
   * @return return true when window is hidden, false when window is visible
   */
  boolean isWindowHidden() {
    return windowHidden;
  }

  /**
   * @param windowHidden true when window is hidden, false when window is visible
   */
  void setWindowHidden(boolean windowHidden) {
    this.windowHidden = windowHidden;
  }

  /**
   * @return true when window is minimized, false when window is open
   */
  boolean isWindowMinimized() {
    return windowMinimized;
  }

  /**
   * @param windowMinimized true when window is minimized, false when window is open
   */
  void setWindowMinimized(boolean windowMinimized) {
    this.windowMinimized = windowMinimized;
  }

  /**
   * @return true when horizontal scroll bar is visible, false when horizontal scroll bar hidden
   */
  boolean isHorizontalScrollBarVisible() {
    return horizontalScrollBarVisible;
  }

  /**
   * @param horizontalScrollBarVisible true when horizontal scroll bar is visible, false when horizontal scroll bar hidden
   */
  void setHorizontalScrollBarVisible(boolean horizontalScrollBarVisible) {
    this.horizontalScrollBarVisible = horizontalScrollBarVisible;
  }

  /**
   * @return true when vertical scroll bar visible, false when vertical scroll bar hidden
   */
  boolean isVerticalScrollBarVisible() {
    return verticalScrollBarVisible;
  }

  /**
   * @param verticalScrollBarVisible true when vertical scroll bar visible, false when vertical scroll bar hidden
   */
  void setVerticalScrollBarVisible(boolean verticalScrollBarVisible) {
    this.verticalScrollBarVisible = verticalScrollBarVisible;
  }

  /**
   * @return true when worksheet tab bar visible, false when worksheet tab bar hidden
   */
  boolean isWorksheetTabBarVisible() {
    return worksheetTabBarVisible;
  }

  /**
   * @param worksheetTabBarVisible true when worksheet tab bar visible, false when worksheet tab bar hidden
   */
  void setWorksheetTabBarVisible(boolean worksheetTabBarVisible) {
    this.worksheetTabBarVisible = worksheetTabBarVisible;
  }

  /**
   * @return the index to active (displayed) worksheet
   */
  int getSelectedSheet() {
    return selectedSheet;
  }

  /**
   * @param selectedSheet Index to active (displayed) worksheet
   */
  void setSelectedSheet(int selectedSheet) {
    this.selectedSheet = selectedSheet;
  }

  /**
   * @return the index of first visible tab in the worksheet tab bar
   */
  int getFirstVisibleTab() {
    return firstVisibleTab;
  }

  /**
   * @param firstVisibleTab Index of first visible tab in the worksheet tab bar
   */
  void setFirstVisibleTab(int firstVisibleTab) {
    this.firstVisibleTab = firstVisibleTab;
  }

  /**
   * @return the number of selected worksheets (highlighted in the worksheet tab bar)
   */
  int getNumberOfSelectedWorkSheets() {
    return numberOfSelectedWorkSheets;
  }

  /**
   * @param numberOfSelectedWorkSheets Number of selected worksheets (highlighted in the worksheet tab bar)
   */
  void setNumberOfSelectedWorkSheets(int numberOfSelectedWorkSheets) {
    this.numberOfSelectedWorkSheets = numberOfSelectedWorkSheets;
  }

  /**
   * @return Width of worksheet tab bar (in 1/1000 of window width).
   */
  int getWidthOfWorksheetTabBar() {
    return widthOfWorksheetTabBar;
  }

  /**
   * @param widthOfWorksheetTabBar Width of worksheet tab bar (in 1/1000 of window width). The remaining space is used by the horizontal scrollbar.
   */
  void setWidthOfWorksheetTabBar(int widthOfWorksheetTabBar) {
    this.widthOfWorksheetTabBar = widthOfWorksheetTabBar;
  }

}
