package jxl.biff.drawing;

import java.io.*;
import jxl.biff.*;
import jxl.write.biff.*;
import jxl.write.biff.File;

/**
 * created 29.05.2020
 * @author jan
 */
public interface Comment extends DrawingGroupObject {

  /**
   * Adds an mso record to this object
   *
   * @param d the mso record
   */
  void addMso(MsoDrawingRecord d);

  /**
   * Accessor for the blip id
   *
   * @return the blip id
   */
  int getBlipId();

  /**
   * Accessor for the column
   *
   * @return  the column
   */
  int getColumn();

  /**
   * Accessor for the drawing group
   *
   * @return the drawing group
   */
  DrawingGroup getDrawingGroup();

  /**
   * Accessor for the height of this drawing
   *
   * @return the number of rows spanned by this image
   */
  double getHeight();

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  byte[] getImageBytes();

  /**
   * Accessor for the image data
   *
   * @return the image data
   */
  byte[] getImageData();

  /**
   * Accessor for the image file path.  Normally this is the absolute path
   * of a file on the directory system, but if this drawing was constructed
   * using an byte[] then the blip id is returned
   *
   * @return the image file path, or the blip id
   */
  String getImageFilePath();

  /**
   * Gets the drawing record which was read in
   *
   * @return the drawing record
   */
  MsoDrawingRecord getMsoDrawingRecord();

  /**
   * Accessor for the object id
   *
   * @return the object id
   */
  int getObjectId();

  /**
   * Gets the origin of this drawing
   *
   * @return where this drawing came from
   */
  Origin getOrigin();

  /**
   * Accessor for the reference count on this drawing
   *
   * @return the reference count
   */
  int getReferenceCount();

  /**
   * Accessor for the row
   *
   * @return  the row
   */
  int getRow();

  /**
   * Accessor for the shape id
   *
   * @return the object id
   */
  @Override
  int getShapeId();

  /**
   * Creates the main Sp container for the drawing
   *
   * @return the SP container
   */
  @Override
  EscherContainer getSpContainer();

  /**
   * Accessor for the comment text
   *
   * @return  the comment text
   */
  String getText();

  /**
   * Accessor for the type
   *
   * @return the type
   */
  @Override
  ShapeType getType();

  /**
   * Accessor for the width of this drawing
   *
   * @return the number of columns spanned by this image
   */
  @Override
  double getWidth();

  /**
   * Accessor for the column of this drawing
   *
   * @return the column
   */
  @Override
  double getX();

  /**
   * Accessor for the row of this drawing
   *
   * @return the row
   */
  @Override
  double getY();

  /**
   * Accessor for the first drawing on the sheet.  This is used when
   * copying unmodified sheets to indicate that this drawing contains
   * the first time Escher gubbins
   *
   * @return TRUE if this MSORecord is the first drawing on the sheet
   */
  @Override
  boolean isFirst();

  /**
   * Queries whether this object is a form object.  Form objects have their
   * drawings records spread over several records and require special handling
   *
   * @return TRUE if this is a form object, FALSE otherwise
   */
  @Override
  boolean isFormObject();

  /**
   * Called when the comment text is changed during the sheet copy process
   *
   * @param t the new text
   */
  void setCommentText(String t);

  /**
   * Sets the drawing group for this drawing.  Called by the drawing group
   * when this drawing is added to it
   *
   * @param dg the drawing group
   */
  @Override
  void setDrawingGroup(DrawingGroup dg);

  /**
   * Sets the formatting
   *
   * @param t the formatting record
   */
  void setFormatting(ContinueRecord t);

  /**
   * Accessor for the height of this drawing
   *
   * @param h the number of rows spanned by this image
   */
  @Override
  void setHeight(double h);

  /**
   * Sets the note object
   *
   * @param t the note record
   */
  void setNote(NoteRecord t);

  /**
   * Sets the object id.  Invoked by the drawing group when the object is
   * added to id
   *
   * @param objid the object id
   * @param bip the blip id
   * @param sid the shape id
   */
  @Override
  void setObjectId(int objid, int bip, int sid);

  /**
   * Sets the new reference count on the drawing
   *
   * @param r the new reference count
   */
  @Override
  void setReferenceCount(int r);

  /**
   * Sets the text data
   *
   * @param t the text data
   */
  void setText(ContinueRecord t);

  /**
   * Sets the text object
   *
   * @param t the text object
   */
  void setTextObject(TextObjectRecord t);

  /**
   * Accessor for the width
   *
   * @param w the number of columns to span
   */
  @Override
  void setWidth(double w);

  /**
   * Sets the column position of this drawing.  Used when inserting/removing
   * columns from the spreadsheet
   *
   * @param x the column
   */
  @Override
  void setX(double x);

  /**
   * Accessor for the row of the drawing
   *
   * @param y the row
   */
  @Override
  void setY(double y);

  /**
   * Writes out the additional comment records
   *
   * @param outputFile the output file
   * @exception IOException
   */
  @Override
  void writeAdditionalRecords(File outputFile) throws IOException;

  /**
   * Writes any records that need to be written after all the drawing group
   * objects have been written
   * Writes out all the note records
   *
   * @param outputFile the output file
   * @exception IOException
   */
  @Override
  void writeTailRecords(File outputFile) throws IOException;

}
