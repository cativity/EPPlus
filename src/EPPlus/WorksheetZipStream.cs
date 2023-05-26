/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using EPPlusTest.Utils;
using OfficeOpenXml.Utils;
using System;
using System.IO;

namespace OfficeOpenXml;

internal class WorksheetZipStream : Stream
{
    RollingBuffer _rollingBuffer = new RollingBuffer(8192);
    private Stream _stream;
    //private long _size;
    //private long _bytesRead;
    //private int _bufferEnd = 0;
    //private int _prevBufferEnd = 0;
    public WorksheetZipStream(Stream zip, bool writeToBuffer, long size = -1)
    {
        this._stream = zip;
        //_size = size;
        //_bytesRead = 0;
        this.WriteToBuffer = writeToBuffer;
    }

    public override bool CanRead => this._stream.CanRead;

    public override bool CanSeek => this._stream.CanSeek;

    public override bool CanWrite => this._stream.CanWrite;

    public override long Length => this._stream.Length;

    public override long Position { get => this._stream.Position; set => this._stream.Position = value; }

    public override void Flush()
    {
        this._stream.Flush();
    }

    //byte[] _buffer = null;
    //byte[] _prevBuffer, _tempBuffer = new byte[8192];
    public override int Read(byte[] buffer, int offset, int count)
    {
        if(this._stream.Length > 0 && this._stream.Position + count > this._stream.Length)
        {
            count = (int)(this._stream.Length - this._stream.Position);
        }

        int r = this._stream.Read(buffer, offset, count);
        if (r > 0)
        {
            if (this.WriteToBuffer)
            {
                this.Buffer.Write(buffer, 0, r);
            }

            this._rollingBuffer.Write(buffer, r);
        }
        return r;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        return this._stream.Seek(offset, origin);
    }

    public override void SetLength(long value)
    {
        this._stream.SetLength(value);
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        this._stream.Write(buffer, offset, count);
    }
    public BinaryWriter Buffer = new BinaryWriter(RecyclableMemory.GetStream());
    public void SetWriteToBuffer()
    {
        this.Buffer = new BinaryWriter(RecyclableMemory.GetStream());
        this.Buffer.Write(this._rollingBuffer.GetBuffer());
        this.WriteToBuffer = true;
    }
    public bool WriteToBuffer { get; set; }

    internal string GetBufferAsString(bool writeToBufferAfter)
    {
        this.WriteToBuffer = writeToBufferAfter;
        this.Buffer.Flush();
        return System.Text.Encoding.UTF8.GetString(((MemoryStream)this.Buffer.BaseStream).ToArray());
    }
    internal string GetBufferAsStringRemovingElement(bool writeToBufferAfter, string element)
    {
        string xml;
        if (this.WriteToBuffer)
        {
            this.Buffer.Flush();
            xml = System.Text.Encoding.UTF8.GetString(((MemoryStream)this.Buffer.BaseStream).ToArray());
        }
        else
        {
            xml = System.Text.Encoding.UTF8.GetString(this._rollingBuffer.GetBuffer());
        }

        this.WriteToBuffer = writeToBufferAfter;
        GetElementPos(xml, element, out int startIx, out int endIx);
        if (startIx > 0)
        {
            return xml.Substring(0, startIx) + GetPlaceholderTag(xml, startIx, endIx);
        }
        else
        {
            return xml;
        }
    }

    private static string GetPlaceholderTag(string xml, int startIx, int endIx)
    {
        string? placeholderTag = xml.Substring(startIx, endIx - startIx);
        placeholderTag = placeholderTag.Replace("/", "");
        placeholderTag = placeholderTag.Substring(0, placeholderTag.Length - 1) + "/>";
        return placeholderTag;
    }

    private static int GetEndElementPos(string xml, string element, int endIx)
    {
        int ix = xml.IndexOf("/" + element + ">", endIx);
        if (ix > 0)
        {
            return ix + element.Length + 2;
        }
        return -1;
    }

    private static void GetElementPos(string xml, string element, out int startIx, out int endIx)
    {
        int ix = -1;
        do
        {
            ix = xml.IndexOf(element, ix + 1);
            if (ix > 0 && (xml[ix - 1] == ':' || xml[ix - 1] == '<'))
            {
                startIx = ix;
                if (startIx >= 0 && xml[startIx] != '<')
                {
                    startIx--;
                }
                endIx = ix + element.Length;
                while (endIx < xml.Length && xml[endIx] == ' ')
                {
                    endIx++;
                }
                if (endIx < xml.Length && xml[endIx] == '>')
                {
                    endIx++;
                    return;
                }
                else if (endIx < xml.Length + 1 && xml.Substring(endIx, 2) == "/>")
                {
                    endIx += 2;
                    return;
                }
            }
        }
        while (ix >= 0);
        startIx = endIx = -1;
    }

    internal void ReadToEnd()
    {
        if (this._stream.Position < this._stream.Length)
        {
            int sizeToEnd = (int)(this._stream.Length - this._stream.Position);
            byte[] buffer = new byte[sizeToEnd];
            int r = this._stream.Read(buffer, 0, sizeToEnd);
            this.Buffer.Write(buffer);
        }
    }

    internal string ReadFromEndElement(string endElement, string startXml = "", string readToElement = null, bool writeToBuffer = true, string xmlPrefix = "", bool addEmptyNode = true)
    {
        if (string.IsNullOrEmpty(readToElement) && this._stream.Position < this._stream.Length)
        {
            this.ReadToEnd();
        }

        this.Buffer.Flush();
        string? xml = System.Text.Encoding.UTF8.GetString(((MemoryStream)this.Buffer.BaseStream).ToArray());
        int endElementIx = FindElementPos(xml, endElement, false);

        if (endElementIx < 0)
        {
            return startXml;
        }

        if (string.IsNullOrEmpty(readToElement))
        {
            xml = xml.Substring(endElementIx);
        }
        else
        {
            int toElementIx = FindElementPos(xml, readToElement);
            if (toElementIx >= endElementIx)
            {
                xml = xml.Substring(endElementIx, toElementIx - endElementIx);
                if (addEmptyNode)
                {
                    xml += string.IsNullOrEmpty(xmlPrefix) ? $"<{readToElement}/>" : $"<{xmlPrefix}:{readToElement}/>";
                }
            }
            else
            {
                xml = xml.Substring(endElementIx);
            }
        }

        this.WriteToBuffer = writeToBuffer;
        return startXml + xml;
    }
    internal string ReadToExt(string startXml, string uriValue, ref string lastElement, string lastUri="")
    {
        this.Buffer.Flush();
        string? xml = System.Text.Encoding.UTF8.GetString(((MemoryStream)this.Buffer.BaseStream).ToArray());

        if (lastElement == "ext")
        {
            int lastExtStartIx = GetXmlIndex(xml, lastUri);
            int endExtIx;
            if(lastExtStartIx < 0)
            {
                endExtIx = FindElementPos(xml, "ext", false);
            }
            else
            {
                endExtIx = FindElementPos(xml, "ext", false, lastExtStartIx+4);
            }
            xml = xml.Substring(endExtIx);
        }
        else
        {
            int lastElementIx = FindElementPos(xml, lastElement, false, 0);
            if (lastElementIx < 0)
            {
                throw new InvalidOperationException("Worksheet Xml is invalid");
            }
            xml = xml.Substring(lastElementIx);
        }
        if (string.IsNullOrEmpty(uriValue))
        {
            lastElement = "";
            return startXml + xml;
        }
        else
        {
            int ix = GetXmlIndex(xml, uriValue);
            if (ix > 0)
            {
                lastElement = "ext";
                return startXml + xml.Substring(0, ix);
            }
        }
        return startXml;
    }

    private static int GetXmlIndex(string xml, string uriValue)
    {
        int elementStartIx = FindElementPos(xml, "ext", true, 0);
        while (elementStartIx > 0)
        {
            int elementEndIx = xml.IndexOf('>', elementStartIx);
            string? elementString = xml.Substring(elementStartIx, elementEndIx - elementStartIx + 1);
            if (HasExtElementUri(elementString, uriValue))
            {
                return elementStartIx;
            }
            elementStartIx = FindElementPos(xml, "ext", true, elementEndIx + 1);
        }
        return -1;
    }

    private static bool HasExtElementUri(string elementString, string uriValue)
    {
        if (elementString.StartsWith("</"))
        {
            return false; //An endtag, return false;
        }

        int ix=elementString.IndexOf("uri");
        char pc = elementString[ix - 1];
        char nc = elementString[ix + 3];
        if(char.IsWhiteSpace(pc) && (char.IsWhiteSpace(nc) || nc=='='))
        {
            ix = elementString.IndexOf('=', ix + 1);
            int ixAttrStart = elementString.IndexOf('"', ix + 1) + 1;
            int ixAttrEnd = elementString.IndexOf('"', ixAttrStart + 1) - 1;

            string? uri = elementString.Substring(ixAttrStart, ixAttrEnd - ixAttrStart+1);
            return uriValue.Equals(uri, StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    /// <summary>
    /// Returns the position in the xml document for an element. Either returns the position of the start element or the end element.
    /// </summary>
    /// <param name="xml">The xml to search</param>
    /// <param name="element">The element</param>
    /// <param name="returnStartPos">If the position before the start element is returned. If false the end of the end element is returned.</param>
    /// <returns>The position of the element in the input xml</returns>
    private static int FindElementPos(string xml, string element, bool returnStartPos = true, int ix=0)
    {
        while (true)
        {
            ix = xml.IndexOf(element, ix);
            if (ix > 0 && ix < xml.Length - 1)
            {
                char c = xml[ix + element.Length];
                if (c == '>' || c == ' ' || c == '/')
                {
                    c = xml[ix - 1];
                    if (c == '/' || c == ':' || xml[ix - 1] == '<')
                    {
                        if (returnStartPos)
                        {
                            return xml.LastIndexOf('<', ix);
                        }
                        else
                        {
                            //Return the end element, either </element> or <element/>
                            int startIx = xml.LastIndexOf("<", ix);
                            if (ix > 0)
                            {
                                int end = xml.IndexOf(">", ix + element.Length - 1);
                                if (xml[startIx + 1] == '/' || xml[end - 1] == '/')
                                {
                                    return end + 1;
                                }
                            }
                        }
                    }
                }
            }
            if (ix < 0)
            {
                return -1;
            }

            ix += element.Length;
        }
    }
}