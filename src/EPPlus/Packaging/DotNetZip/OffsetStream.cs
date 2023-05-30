// OffsetStream.cs
// ------------------------------------------------------------------
//
// Copyright (c)  2009 Dino Chiesa
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
// Time-stamp: <2009-August-27 12:50:35>
//
// ------------------------------------------------------------------
//
// This module defines logic for handling reading of zip archives embedded 
// into larger streams.  The initial position of the stream serves as
// the base offset for all future Seek() operations.
// 
// ------------------------------------------------------------------

using System;
using System.IO;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

internal class OffsetStream : Stream, IDisposable
{
    private long _originalPosition;
    private Stream _innerStream;

    public OffsetStream(Stream s)
        : base()
    {
        this._originalPosition = s.Position;
        this._innerStream = s;
    }

    public override int Read(byte[] buffer, int offset, int count) => this._innerStream.Read(buffer, offset, count);

    public override void Write(byte[] buffer, int offset, int count) => throw new NotImplementedException();

    public override bool CanRead => this._innerStream.CanRead;

    public override bool CanSeek => this._innerStream.CanSeek;

    public override bool CanWrite => false;

    public override void Flush() => this._innerStream.Flush();

    public override long Length => this._innerStream.Length;

    public override long Position
    {
        get => this._innerStream.Position - this._originalPosition;
        set => this._innerStream.Position = this._originalPosition + value;
    }

    public override long Seek(long offset, SeekOrigin origin) => this._innerStream.Seek(this._originalPosition + offset, origin) - this._originalPosition;

    public override void SetLength(long value) => throw new NotImplementedException();

    void IDisposable.Dispose() => this.Close();

    public override void Close() => base.Close();
}