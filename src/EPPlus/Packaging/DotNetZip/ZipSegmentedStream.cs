// ZipSegmentedStream.cs
// ------------------------------------------------------------------
//
// Copyright (c) 2009-2011 Dino Chiesa.
// All rights reserved.
//
// This code module is part of DotNetZip, a zipfile class library.
//
// ------------------------------------------------------------------
//
// This code is licensed under the Microsoft Public License.
// See the file License.txt for the license details.
// More info on: http://dotnetzip.codeplex.com
//
// ------------------------------------------------------------------
//
// last saved (in emacs):
// Time-stamp: <2011-July-13 22:25:45>
//
// ------------------------------------------------------------------
//
// This module defines logic for zip streams that span disk files.
//
// ------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

internal class ZipSegmentedStream : Stream
{
    enum RwMode
    {
        None = 0,
        ReadOnly = 1,
        Write = 2,

        //Update = 3
    }

    private RwMode rwMode;
    private bool _exceptionPending; // **see note below
    private string _baseName;

    private string _baseDir;

    //private bool _isDisposed;
    private string _currentName;
    private string _currentTempName;
    private uint _currentDiskNumber;
    private uint _maxDiskNumber;
    private int _maxSegmentSize;
    private Stream _innerStream;

    // **Note regarding exceptions:
    //
    // When ZipSegmentedStream is employed within a using clause,
    // which is the typical scenario, and an exception is thrown
    // within the scope of the using, Dispose() is invoked
    // implicitly before processing the initial exception.  If that
    // happens, this class sets _exceptionPending to true, and then
    // within the Dispose(bool), takes special action as
    // appropriate. Need to be careful: any additional exceptions
    // will mask the original one.

    private ZipSegmentedStream()
        : base()
    {
        this._exceptionPending = false;
    }

    public static ZipSegmentedStream ForReading(string name, uint initialDiskNumber, uint maxDiskNumber)
    {
        ZipSegmentedStream zss = new ZipSegmentedStream()
        {
            rwMode = RwMode.ReadOnly, CurrentSegment = initialDiskNumber, _maxDiskNumber = maxDiskNumber, _baseName = name,
        };

        // Console.WriteLine("ZSS: ForReading ({0})",
        //                    Path.GetFileName(zss.CurrentName));

        zss._SetReadStream();

        return zss;
    }

    public static ZipSegmentedStream ForWriting(string name, int maxSegmentSize)
    {
        ZipSegmentedStream zss = new ZipSegmentedStream()
        {
            rwMode = RwMode.Write, CurrentSegment = 0, _baseName = name, _maxSegmentSize = maxSegmentSize, _baseDir = Path.GetDirectoryName(name)
        };

        // workitem 9522
        if (zss._baseDir == "")
        {
            zss._baseDir = ".";
        }

        zss._SetWriteStream(0);

        // Console.WriteLine("ZSS: ForWriting ({0})",
        //                    Path.GetFileName(zss.CurrentName));

        return zss;
    }

    /// <summary>
    ///   Sort-of like a factory method, ForUpdate is used only when
    ///   the application needs to update the zip entry metadata for
    ///   a segmented zip file, when the starting segment is earlier
    ///   than the ending segment, for a particular entry.
    /// </summary>
    /// <remarks>
    ///   <para>
    ///     The update is always contiguous, never rolls over.  As a
    ///     result, this method doesn't need to return a ZSS; it can
    ///     simply return a FileStream.  That's why it's "sort of"
    ///     like a Factory method.
    ///   </para>
    ///   <para>
    ///     Caller must Close/Dispose the stream object returned by
    ///     this method.
    ///   </para>
    /// </remarks>
    public static Stream ForUpdate(string name, uint diskNumber)
    {
        if (diskNumber >= 99)
        {
            throw new ArgumentOutOfRangeException("diskNumber");
        }

        string fname = String.Format("{0}.z{1:D2}", Path.Combine(Path.GetDirectoryName(name), Path.GetFileNameWithoutExtension(name)), diskNumber + 1);

        // Console.WriteLine("ZSS: ForUpdate ({0})",
        //                   Path.GetFileName(fname));

        // This class assumes that the update will not expand the
        // size of the segment. Update is used only for an in-place
        // update of zip metadata. It never will try to write beyond
        // the end of a segment.

        return File.Open(fname, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }

    public bool ContiguousWrite { get; set; }

    public UInt32 CurrentSegment
    {
        get { return this._currentDiskNumber; }
        private set
        {
            this._currentDiskNumber = value;
            this._currentName = null; // it will get updated next time referenced
        }
    }

    /// <summary>
    ///   Name of the filesystem file corresponding to the current segment.
    /// </summary>
    /// <remarks>
    ///   <para>
    ///     The name is not always the name currently being used in the
    ///     filesystem.  When rwMode is RwMode.Write, the filesystem file has a
    ///     temporary name until the stream is closed or until the next segment is
    ///     started.
    ///   </para>
    /// </remarks>
    public String CurrentName
    {
        get { return this._currentName ??= this._NameForSegment(this.CurrentSegment); }
    }

    public String CurrentTempName
    {
        get { return this._currentTempName; }
    }

    private string _NameForSegment(uint diskNumber)
    {
        if (diskNumber >= 99)
        {
            this._exceptionPending = true;

            throw new OverflowException("The number of zip segments would exceed 99.");
        }

        return String.Format("{0}.z{1:D2}",
                             Path.Combine(Path.GetDirectoryName(this._baseName), Path.GetFileNameWithoutExtension(this._baseName)),
                             diskNumber + 1);
    }

    // Returns the segment that WILL be current if writing
    // a block of the given length.
    // This isn't exactly true. It could roll over beyond
    // this number.
    public UInt32 ComputeSegment(int length)
    {
        if (this._innerStream.Position + length > this._maxSegmentSize)

            // the block will go AT LEAST into the next segment
        {
            return this.CurrentSegment + 1;
        }

        // it will fit in the current segment
        return this.CurrentSegment;
    }

    public override String ToString()
    {
        return String.Format("{0}[{1}][{2}], pos=0x{3:X})", "ZipSegmentedStream", this.CurrentName, this.rwMode.ToString(), this.Position);
    }

    private void _SetReadStream()
    {
        if (this._innerStream != null)
        {
#if NETCF
                _innerStream.Close();
#else
            this._innerStream.Dispose();
#endif
        }

        if (this.CurrentSegment + 1 == this._maxDiskNumber)
        {
            this._currentName = this._baseName;
        }

        // Console.WriteLine("ZSS: SRS ({0})",
        //                   Path.GetFileName(CurrentName));

        this._innerStream = File.OpenRead(this.CurrentName);
    }

    /// <summary>
    /// Read from the stream
    /// </summary>
    /// <param name="buffer">the buffer to read</param>
    /// <param name="offset">the offset at which to start</param>
    /// <param name="count">the number of bytes to read</param>
    /// <returns>the number of bytes actually read</returns>
    public override int Read(byte[] buffer, int offset, int count)
    {
        if (this.rwMode != RwMode.ReadOnly)
        {
            this._exceptionPending = true;

            throw new InvalidOperationException("Stream Error: Cannot Read.");
        }

        int r = this._innerStream.Read(buffer, offset, count);
        int r1 = r;

        while (r1 != count)
        {
            if (this._innerStream.Position != this._innerStream.Length)
            {
                this._exceptionPending = true;

                throw new ZipException(String.Format("Read error in file {0}", this.CurrentName));
            }

            if (this.CurrentSegment + 1 == this._maxDiskNumber)
            {
                return r; // no more to read
            }

            this.CurrentSegment++;
            this._SetReadStream();
            offset += r1;
            count -= r1;
            r1 = this._innerStream.Read(buffer, offset, count);
            r += r1;
        }

        return r;
    }

    private void _SetWriteStream(uint increment)
    {
        if (this._innerStream != null)
        {
#if NETCF
                _innerStream.Close();
#else
            this._innerStream.Dispose();
#endif
            if (File.Exists(this.CurrentName))
            {
                File.Delete(this.CurrentName);
            }

            File.Move(this._currentTempName, this.CurrentName);

            // Console.WriteLine("ZSS: SWS close ({0})",
            //                   Path.GetFileName(CurrentName));
        }

        if (increment > 0)
        {
            this.CurrentSegment += increment;
        }

        SharedUtilities.CreateAndOpenUniqueTempFile(this._baseDir, out this._innerStream, out this._currentTempName);

        // Console.WriteLine("ZSS: SWS open ({0})",
        //                   Path.GetFileName(_currentTempName));

        if (this.CurrentSegment == 0)
        {
            this._innerStream.Write(BitConverter.GetBytes(ZipConstants.SplitArchiveSignature), 0, 4);
        }
    }

    /// <summary>
    /// Write to the stream.
    /// </summary>
    /// <param name="buffer">the buffer from which to write</param>
    /// <param name="offset">the offset at which to start writing</param>
    /// <param name="count">the number of bytes to write</param>
    public override void Write(byte[] buffer, int offset, int count)
    {
        if (this.rwMode != RwMode.Write)
        {
            this._exceptionPending = true;

            throw new InvalidOperationException("Stream Error: Cannot Write.");
        }

        if (this.ContiguousWrite)
        {
            // enough space for a contiguous write?
            if (this._innerStream.Position + count > this._maxSegmentSize)
            {
                this._SetWriteStream(1);
            }
        }
        else
        {
            while (this._innerStream.Position + count > this._maxSegmentSize)
            {
                int c = unchecked(this._maxSegmentSize - (int)this._innerStream.Position);
                this._innerStream.Write(buffer, offset, c);
                this._SetWriteStream(1);
                count -= c;
                offset += c;
            }
        }

        this._innerStream.Write(buffer, offset, count);
    }

    public long TruncateBackward(uint diskNumber, long offset)
    {
        // Console.WriteLine("***ZSS.Trunc to disk {0}", diskNumber);
        // Console.WriteLine("***ZSS.Trunc:  current disk {0}", CurrentSegment);
        if (diskNumber >= 99)
        {
            throw new ArgumentOutOfRangeException("diskNumber");
        }

        if (this.rwMode != RwMode.Write)
        {
            this._exceptionPending = true;

            throw new ZipException("bad state.");
        }

        // Seek back in the segmented stream to a (maybe) prior segment.

        // Check if it is the same segment.  If it is, very simple.
        if (diskNumber == this.CurrentSegment)
        {
            long x = this._innerStream.Seek(offset, SeekOrigin.Begin);

            // workitem 10178
            SharedUtilities.Workaround_Ladybug318918(this._innerStream);

            return x;
        }

        // Seeking back to a prior segment.
        // The current segment and any intervening segments must be removed.
        // First, close the current segment, and then remove it.
        if (this._innerStream != null)
        {
#if NETCF
                _innerStream.Close();
#else
            this._innerStream.Dispose();
#endif
            if (File.Exists(this._currentTempName))
            {
                File.Delete(this._currentTempName);
            }
        }

        // Now, remove intervening segments.
        for (uint j = this.CurrentSegment - 1; j > diskNumber; j--)
        {
            string s = this._NameForSegment(j);

            // Console.WriteLine("***ZSS.Trunc:  removing file {0}", s);
            if (File.Exists(s))
            {
                File.Delete(s);
            }
        }

        // now, open the desired segment.  It must exist.
        this.CurrentSegment = diskNumber;

        // get a new temp file, try 3 times:
        for (int i = 0; i < 3; i++)
        {
            try
            {
                this._currentTempName = SharedUtilities.InternalGetTempFileName();

                // move the .z0x file back to a temp name
                File.Move(this.CurrentName, this._currentTempName);

                break; // workitem 12403
            }
            catch (IOException)
            {
                if (i == 2)
                {
                    throw;
                }
            }
        }

        // open it
        this._innerStream = new FileStream(this._currentTempName, FileMode.Open);

        long r = this._innerStream.Seek(offset, SeekOrigin.Begin);

        // workitem 10178
        SharedUtilities.Workaround_Ladybug318918(this._innerStream);

        return r;
    }

    public override bool CanRead
    {
        get { return this.rwMode == RwMode.ReadOnly && this._innerStream != null && this._innerStream.CanRead; }
    }

    public override bool CanSeek
    {
        get { return this._innerStream != null && this._innerStream.CanSeek; }
    }

    public override bool CanWrite
    {
        get { return this.rwMode == RwMode.Write && this._innerStream != null && this._innerStream.CanWrite; }
    }

    public override void Flush()
    {
        this._innerStream.Flush();
    }

    public override long Length
    {
        get { return this._innerStream.Length; }
    }

    public override long Position
    {
        get { return this._innerStream.Position; }
        set { this._innerStream.Position = value; }
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        long x = this._innerStream.Seek(offset, origin);

        // workitem 10178
        SharedUtilities.Workaround_Ladybug318918(this._innerStream);

        return x;
    }

    public override void SetLength(long value)
    {
        if (this.rwMode != RwMode.Write)
        {
            this._exceptionPending = true;

            throw new InvalidOperationException();
        }

        this._innerStream.SetLength(value);
    }

    protected override void Dispose(bool disposing)
    {
        // this gets called by Stream.Close()

        // if (_isDisposed) return;
        // _isDisposed = true;
        //Console.WriteLine("Dispose (mode={0})\n", rwMode.ToString());

        try
        {
            if (this._innerStream != null)
            {
#if NETCF
                    _innerStream.Close();
#else
                this._innerStream.Dispose();
#endif

                //_innerStream = null;
                if (this.rwMode == RwMode.Write)
                {
                    if (this._exceptionPending)
                    {
                        // possibly could try to clean up all the
                        // temp files created so far...
                    }
                    else
                    {
                        // // move the final temp file to the .zNN name
                        // if (File.Exists(CurrentName))
                        //     File.Delete(CurrentName);
                        // if (File.Exists(_currentTempName))
                        //     File.Move(_currentTempName, CurrentName);
                    }
                }
            }
        }
        finally
        {
            base.Dispose(disposing);
        }
    }
}