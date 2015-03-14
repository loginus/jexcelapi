/*********************************************************************
*
*      Copyright (C) 2007 Andrew Khan
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

import java.io.*;
import java.nio.file.*;
import jxl.common.Logger;

/**
 * Used to generate the excel biff data using a temporary file.  This
 * class wraps a RandomAccessFile
 */
class FileDataOutput implements ExcelDataOutput
{

  /** 
   * The temporary file
   */
  private final Path temporaryFile;

  /**
   * The excel data
   */
  private final RandomAccessFile data;

  /**
   * Constructor
   *
   * @param tmpdir the temporary directory used to write files.  If this is
   *               NULL then the sytem temporary directory will be used
   */
  public FileDataOutput(Path tmpdir) throws IOException
  {
    temporaryFile = Files.createTempFile(tmpdir,"jxl",".tmp");
    temporaryFile.toFile().deleteOnExit();
    data = new RandomAccessFile(temporaryFile.toFile(), "rw");
  }

  /**
   * Writes the bytes to the end of the array, growing the array
   * as needs dictate
   *
   * @param d the data to write to the end of the array
   */
  public void write(byte[] bytes) throws IOException
  {
    data.write(bytes);
  }

  /**
   * Gets the current position within the file
   *
   * @return the position within the file
   */
  @Override
  public int getPosition() throws IOException
  {
    // As all excel data structures are four bytes anyway, it's ok to 
    // truncate the long to an int
    return (int) data.getFilePointer();
  }

  /**
   * Sets the data at the specified position to the contents of the array
   * 
   * @param pos the position to alter
   * @param newdata the data to modify
   */
  public void setData(byte[] newdata, int pos) throws IOException
  {
    long curpos = data.getFilePointer();
    data.seek(pos);
    data.write(newdata);
    data.seek(curpos);
  }

  /** 
   * Writes the data to the output stream
   */
  @Override
  public void writeData(OutputStream out) throws IOException
  {
		byte[] buffer = new byte[1024];
		int length = 0;
		data.seek(0);
		while ((length = data.read(buffer)) != -1)
		{
			out.write(buffer, 0, length);
		}
  }

  /**
   * Called when the final compound file has been written
   */
  @Override
  public void close() throws IOException
  {
    data.close();

    // Explicitly delete the temporary file, since sometimes it is the case
    // that a single process may be generating multiple different excel files
    Files.delete(temporaryFile);
  }
}
