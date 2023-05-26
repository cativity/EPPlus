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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace OfficeOpenXml.Utils.CompundDocument;

/// <summary>
/// Reads and writes a compound documents.
/// Read spec here https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/[MS-CFB].pdf
/// </summary>
internal partial class CompoundDocumentFile : IDisposable
{
    internal CompoundDocumentFile()
    {
        this.RootItem = new CompoundDocumentItem() { Name = "<Root>", Children=new List<CompoundDocumentItem>(), ObjectType=5 };
        this.minorVersion = 0x3E;
        this.majorVersion = 3;
        this.sectorShif = 9;
        this.minSectorShift = 6;

        this._sectorSize = 1 << this.sectorShif;
        this._miniSectorSize = 1 << this.minSectorShift;
        this._sectorSizeInt = this._sectorSize / 4;
    }
    internal CompoundDocumentFile(FileInfo fi) : this(File.ReadAllBytes(fi.FullName))
    {
            
    }
    internal CompoundDocumentFile(byte[] file)
    {
        using MemoryStream? ms = RecyclableMemory.GetStream(file);
        this.LoadFromMemoryStream(ms);
    }
    internal CompoundDocumentFile(MemoryStream ms)
    {
        this.LoadFromMemoryStream(ms);
    }
    private struct DocWriteInfo
    {
        internal List<int> DIFAT, FAT, miniFAT;
    }
    #region Constants
    const int miniFATSectorSize = 64;
    const int FATSectorSizeV3= 512;
    const int FATSectorSizeV4 = 4096;

    const int DIFAT_SECTOR = -4; //0xFFFFFFFC;
    const int FAT_SECTOR = -3;   //0xFFFFFFFD;
    const int END_OF_CHAIN = -2; //0xFFFFFFFE;
    const int FREE_SECTOR = -1;  //0xFFFFFFFF;

    static readonly byte[] header = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
    #endregion
    #region Private Fields
    short minorVersion;
    short majorVersion;            
    int numberOfDirectorySector;
    short sectorShif, minSectorShift;       //Bits for sector size

    int _numberOfFATSectors;            // (4 bytes): This integer field contains the count of the number of FAT sectors in the compound file.
    int _firstDirectorySectorLocation;  // (4 bytes): This integer field contains the starting sector number for the directory stream.            
    int _transactionSignatureNumber;    // (4 bytes): This integer field MAY contain a sequence number that is incremented every time the compound file is saved by an implementation that supports file transactions.This is the field that MUST be set to all zeroes if file transactions are not implemented.<1> 
    int _miniStreamCutoffSize;          // (4 bytes): This integer field MUST be set to 0x00001000. This field specifies the maximum size of a user-defined data stream that is allocated from the mini FAT and mini stream, and that cutoff is 4,096 bytes.Any user-defined data stream that is larger than or equal to this cutoff size must be allocated as normal sectors from the FAT.
    int _firstMiniFATSectorLocation;    // (4 bytes): This integer field contains the starting sector number for the mini FAT. 
    int _numberofMiniFATSectors;        // (4 bytes): This integer field contains the count of the number of mini FAT sectors in the compound file. 
    int _firstDIFATSectorLocation;      // (4 bytes): This integer field contains the starting sector number for the DIFAT. 
    int _numberofDIFATSectors;          // (4 bytes): This integer field contains the count of the number of DIFAT sectors in the compound file. 

    List<byte[]> _sectors, _miniSectors;
    int _sectorSize, _miniSectorSize;
    int _sectorSizeInt;
    int _currentDIFATSectorPos, _currentFATSectorPos, _currentDirSectorPos;
    int _prevDirFATSectorPos;

    #endregion
    public CompoundDocumentItem RootItem { get; set; }
    List<CompoundDocumentItem> _directories = null;
    internal List<CompoundDocumentItem> Directories
    {
        get { return this._directories ??= this.FlattenDirs(); }
    }

    /// <summary>
    /// Verifies that the header is correct.
    /// </summary>
    /// <param name="fi">The file</param>
    /// <returns></returns>
    public static bool IsCompoundDocument(FileInfo fi)
    {
        try
        {
            using FileStream? fs = fi.OpenRead();
            byte[]? b = new byte[8];
            fs.Read(b, 0, 8);
            fs.Close();
            return IsCompoundDocument(b);
        }
        catch
        {
            return false;
        }            
    }
    public static bool IsCompoundDocument(MemoryStream ms)
    {
        long pos = ms.Position;
        ms.Position = 0;
        byte[]? b=new byte[8];
        ms.Read(b, 0, 8);
        ms.Position = pos;
        return IsCompoundDocument(b);
    }
    public static bool IsCompoundDocument(byte[] b)
    {
        if (b==null || b.Length < 8)
        {
            return false;
        }

        for (int i = 0; i < 8; i++)
        {
            if (b[i] != header[i])
            {
                return false;
            }
        }
        return true;
    }
    #region Read
    internal void Read(BinaryReader br)
    {
        br.ReadBytes(8);    //Read header
        br.ReadBytes(16);   //Header CLSID (16 bytes): Reserved and unused class ID that MUST be set to all zeroes (CLSID_NULL). 
        this.minorVersion = br.ReadInt16();
        this.majorVersion = br.ReadInt16();
        br.ReadInt16(); //Byte order
        this.sectorShif = br.ReadInt16();
        this.minSectorShift = br.ReadInt16();

        this._sectorSize = 1 << this.sectorShif;
        this._miniSectorSize = 1 << this.minSectorShift;
        this._sectorSizeInt = this._sectorSize / 4;
        br.ReadBytes(6);    //Reserved
        this.numberOfDirectorySector = br.ReadInt32();
        this._numberOfFATSectors = br.ReadInt32();
        this._firstDirectorySectorLocation = br.ReadInt32();
        this._transactionSignatureNumber = br.ReadInt32();
        this._miniStreamCutoffSize = br.ReadInt32();
        this._firstMiniFATSectorLocation = br.ReadInt32();
        this._numberofMiniFATSectors = br.ReadInt32();
        this._firstDIFATSectorLocation = br.ReadInt32();
        this._numberofDIFATSectors = br.ReadInt32();
        DocWriteInfo dwi = new DocWriteInfo() { DIFAT = new List<int>(), FAT = new List<int>(), miniFAT = new List<int>() };

        for (int i = 0; i < 109; i++)
        {
            int d = br.ReadInt32();
            if (d >= 0)
            {   
                dwi.DIFAT.Add(d);
            }
        }

        this.LoadSectors(br);
        if (this._firstDIFATSectorLocation > 0)
        {
            this.LoadDIFATSectors(dwi);
        }

        dwi.FAT = this.ReadFAT(this._sectors, dwi);
        List<CompoundDocumentItem>? dir = this.ReadDirectories(this._sectors, dwi);

        this.LoadMinSectors(ref dwi, dir);
        foreach (CompoundDocumentItem? d in dir)
        {
            if (d.Stream == null && d.StreamSize > 0)
            {
                if (d.StreamSize < this._miniStreamCutoffSize)
                {
                    d.Stream = GetStream(d.StartingSectorLocation, d.StreamSize, dwi.miniFAT, this._miniSectors);
                }
                else
                {
                    d.Stream = GetStream(d.StartingSectorLocation, d.StreamSize, dwi.FAT, this._sectors);
                }
            }
        }

        this.AddChildTree(dir[0], dir);
        this._directories = dir;
    }
    private void LoadDIFATSectors(DocWriteInfo dwi)
    {
        int nextSector = this._firstDIFATSectorLocation;
        while (nextSector > 0)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream(this._sectors[nextSector]);
            BinaryReader? brDI = new BinaryReader(ms);
            int sect = -1;
            while (ms.Position < this._sectorSize)
            {
                if (sect > 0)
                {
                    dwi.DIFAT.Add(sect);
                }
                sect = brDI.ReadInt32();
            }
            nextSector = sect;
        }
    }
    private void LoadSectors(BinaryReader br)
    {
        this._sectors = new List<byte[]>();
        while (br.BaseStream.Position < br.BaseStream.Length)
        {
            this._sectors.Add(br.ReadBytes(this._sectorSize));
        }
    }
    private void LoadMinSectors(ref DocWriteInfo dwi, List<CompoundDocumentItem> dir)
    {
        dwi.miniFAT = this.ReadMiniFAT(this._sectors,dwi);
        dir[0].Stream = GetStream(dir[0].StartingSectorLocation, dir[0].StreamSize, dwi.FAT, this._sectors);
        this.GetMiniSectors(dir[0].Stream);
    }
    private void GetMiniSectors(byte[] miniFATStream)
    {
        using MemoryStream? ms = RecyclableMemory.GetStream(miniFATStream);
        BinaryReader? br = new BinaryReader(ms);
        this._miniSectors = new List<byte[]>();
        while (ms.Position < ms.Length)
        {
            this._miniSectors.Add(br.ReadBytes(this._miniSectorSize));
        }
    }
    private static byte[] GetStream(int startingSectorLocation, long streamSize, List<int> FAT, List<byte[]> sectors)
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter? bw = new BinaryWriter(ms);

        int size = 0;
        int nextSector = startingSectorLocation;
        while (size < streamSize)
        {
            if (streamSize > size + sectors[nextSector].Length)
            {
                bw.Write(sectors[nextSector]);
                size += sectors[nextSector].Length;
            }
            else
            {
                byte[]? part = new byte[streamSize - size];
                Array.Copy(sectors[nextSector], part, (int)streamSize - size);
                bw.Write(part);
                size += part.Length;
            }
            nextSector = FAT[nextSector];
        }
        bw.Flush();
        return ms.ToArray();
    }
    private List<int> ReadMiniFAT(List<byte[]> sectors, DocWriteInfo dwi)
    {
        List<int>? l = new List<int>();
        int nextSector = this._firstMiniFATSectorLocation;
        while(nextSector!=END_OF_CHAIN)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream(sectors[nextSector]);
            BinaryReader? br = new BinaryReader(ms);
            while (ms.Position < this._sectorSize)
            {
                int d = br.ReadInt32();
                l.Add(d);
            }
            nextSector = dwi.FAT[nextSector];
        }
        return l;
    }
    private List<CompoundDocumentItem> ReadDirectories(List<byte[]> sectors, DocWriteInfo dwi)
    {
        List<CompoundDocumentItem>? dir = new List<CompoundDocumentItem>();
        int nextSector = this._firstDirectorySectorLocation;
        while (nextSector != END_OF_CHAIN)
        {
            ReadDirectory(sectors, nextSector, dir);
            nextSector = dwi.FAT[nextSector];
        }
        return dir;
    }
    private List<int> ReadFAT(List<byte[]> sectors, DocWriteInfo dwi)
    {
        List<int>? l = new List<int>();
        foreach (int i in dwi.DIFAT)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream(sectors[i]);
            BinaryReader? br = new BinaryReader(ms);
            while (ms.Position < this._sectorSize)
            {
                int d = br.ReadInt32();
                l.Add(d);
            }
        }
        return l;
    }
    private static void ReadDirectory(List<byte[]> sectors, int index, List<CompoundDocumentItem> l)
    {
        using MemoryStream? ms = RecyclableMemory.GetStream(sectors[index]);
        BinaryReader? br = new BinaryReader(ms);

        while (ms.Position < ms.Length)
        {
            CompoundDocumentItem? e = new CompoundDocumentItem();
            e.Read(br);
            if (e.ObjectType != 0)
            {
                l.Add(e);
            }
        }
    }
    internal void AddChildTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
    {
        if (e._handled == true)
        {
            return;
        }

        e._handled = true;
        if (e.ChildID > 0)
        {
            CompoundDocumentItem? c = dirs[e.ChildID];
            c.Parent = e;
            e.Children.Add(c);
            this.AddChildTree(c, dirs);
        }
        if (e.LeftSibling > 0)
        {
            CompoundDocumentItem? c = dirs[e.LeftSibling];
            c.Parent = e.Parent;
            c.Parent.Children.Insert(e.Parent.Children.IndexOf(e), c);
            this.AddChildTree(c, dirs);
        }
        if (e.RightSibling > 0)
        {
            CompoundDocumentItem? c = dirs[e.RightSibling];
            c.Parent = e.Parent;
            e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, c);
            this.AddChildTree(c, dirs);
        }
        if (e.ObjectType == 5)
        {
            this.RootItem = e;
        }
    }
    internal void AddLeftSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
    {
        if (e.LeftSibling > 0)
        {
            CompoundDocumentItem? c = dirs[e.LeftSibling];
            if (c.Parent != null)
            {
                c.Parent = e.Parent;
                c.Parent.Children.Insert(e.Parent.Children.IndexOf(e), c);
                e._handled = true;
                this.AddLeftSiblingTree(c, dirs);
            }
        }
    }
    internal void AddRightSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
    {
        if (e.RightSibling > 0)
        {
            CompoundDocumentItem? c = dirs[e.RightSibling];
            c.Parent = e.Parent;
            e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, c);
            e._handled = true;
            this.AddRightSiblingTree(c, dirs);
        }
    }
    #endregion
    #region Write
    public void Write(MemoryStream ms)
    {
        BinaryWriter? bw = new BinaryWriter(ms);

        //InitValues
        this.minorVersion = 62;
        this.majorVersion = 3;
        this.sectorShif = 9;                 //2^9=512 bytes for version 3 documents 
        this.minSectorShift = 6;             //2^6=64 bytes
        this._miniStreamCutoffSize = 4096;
        this._transactionSignatureNumber = 0;
        this._firstDIFATSectorLocation = END_OF_CHAIN;
        this._firstDirectorySectorLocation = 1;
        this._firstMiniFATSectorLocation = 2;
        this._numberOfFATSectors = 1;

        this._currentDIFATSectorPos = 76;             //DIFAT Position in the header
        this._currentFATSectorPos = this._sectorSize;      //First FAT sector starts at Sector 0
        this._currentDirSectorPos = this._sectorSize * 2;  //First FAT sector starts at Sector 1
        this._prevDirFATSectorPos = this._sectorSize + 4;  //Dir sector starts FAT position 1 (4 for int size)

        bw.Write(new byte[512 * 4]);            //Allocate for Header and first FAT, Directory och MiniFAT sectors
        this.WritePosition(bw, 0, ref this._currentDIFATSectorPos, false);
        WritePosition(bw, new int[] { FAT_SECTOR, END_OF_CHAIN, END_OF_CHAIN }, ref this._currentFATSectorPos);  //First sector is first FAT sector, second is First Dir sector, thirs is first Mini FAT sector.

        //Write directories
        this.WriteDirs(bw, this.Directories);
                
        //Fill empty DISectors up to 109
        this.FillDIFAT(bw);
        //Finally write the header information in the top of the file
        this.WriteHeader(bw);
    }

    private List<CompoundDocumentItem> FlattenDirs()
    {
        List<CompoundDocumentItem>? l = new List<CompoundDocumentItem>();
        InitItem(this.RootItem);
        l.Add(this.RootItem);
        this.RootItem.ChildID = AddChildren(this.RootItem, l);
        return l;
    }

    private static void InitItem(CompoundDocumentItem item)
    {
        item.LeftSibling = -1;
        item.RightSibling = -1;
        item._handled = false;
    }

    private static int AddChildren(CompoundDocumentItem item, List<CompoundDocumentItem> l)
    {
        int childId = -1;
        item.ColorFlag = 1; //Always Black-No matter here, we just add nodes as a b-tree
        if (item.Children.Count > 0)
        {
            foreach(CompoundDocumentItem? c in item.Children)
            {
                InitItem(c);
            }

            item.Children.Sort();

            childId=SetSiblings(l.Count, item.Children, 0, item.Children.Count-1, -1);
            l.AddRange(item.Children);
            foreach (CompoundDocumentItem? c in item.Children)
            {
                c.ChildID=AddChildren(c, l);
            }
        }
        return childId;
    }

    private static void SetUnhandled(int listAdd, List<CompoundDocumentItem> children)
    {
        for(int i=0;i<children.Count;i++)
        {
            if(children[i]._handled==false)
            {
                if(i>0 && children[i-1].RightSibling==-1 && children[i].LeftSibling!=i+listAdd-1)
                {
                    children[i - 1].RightSibling = i + listAdd;
                }
                else if (i<children.Count-1 && children[i + 1].LeftSibling == -1 && children[i].RightSibling != i + listAdd+1)
                {
                    children[i + 1].LeftSibling = i + listAdd;
                }
                else
                {
                    throw (new InvalidOperationException("Invalid sibling handling in Document"));
                }
            }
        }
    }

    private static int SetSiblings(int listAdd, List<CompoundDocumentItem> children, int fromPos, int toPos, int currSibl)
    {
        int pos = GetPos(fromPos,toPos);

        CompoundDocumentItem? item = children[pos];
        if (item._handled)
        {
            return currSibl;
        }

        item._handled = true;
        if (fromPos == toPos)
        {
            return fromPos + listAdd;
        }

        int div = pos / 2;
        if (div <= 0)
        {
            div = 1;
        }

        int lPos = GetPos(fromPos, pos-1);
        int rPos = GetPos(pos+1, toPos);
        if (div == 1 && children[lPos]._handled && children[rPos]._handled)
        {
            return pos+ listAdd;
        }

        if (lPos>-1 && lPos >= fromPos)
        {
            item.LeftSibling = SetSiblings(listAdd, children, fromPos, pos-1, item.LeftSibling);
        }
        if (rPos < children.Count && rPos <= toPos)
        {
            item.RightSibling = SetSiblings(listAdd, children, pos+1, toPos, item.RightSibling);
        }
        return pos + listAdd;
    }

    private static int GetPos(int fromPos, int toPos)
    {
        int div=(toPos - fromPos) / 2;
        return fromPos + div;
    }

    private static bool NoGreater(List<CompoundDocumentItem> children, int pos, int lPos, int listAdd)
    {
        if (pos - lPos <= 1)
        {
            return true;
        }

        for(int i=lPos+1;i<=pos; i++)
        {
            if (children[i].RightSibling!=-1 && children[i].RightSibling > lPos+ listAdd)
            {
                return false;
            }
        }
        return true;
    }
    private static bool NoLess(List<CompoundDocumentItem> children, int pos, int rPos, int listAdd)
    {
        if (rPos - pos <= 1)
        {
            return true;
        }

        for (int i = pos + 1; i <= rPos; i++)
        {
            if (children[i].LeftSibling != -1 && children[i].LeftSibling < rPos+ listAdd)
            {
                return false;
            }
        }
        return true;
    }

    private static int GetLevels(int c)
    {
        c--;
        int i = 0;
        while(c>0)
        {
            c >>=  1;
            i++;
        }
        return i;
    }

    private void FillDIFAT(BinaryWriter bw)
    {
        if (this._currentDIFATSectorPos < this._sectorSize)
        {
            bw.Seek(this._currentDIFATSectorPos, SeekOrigin.Begin);
            while (this._currentDIFATSectorPos < this._sectorSize)
            {
                if (this._currentDIFATSectorPos < 512)
                {
                    bw.Write(0xFFFFFFFF);
                }
                else
                {
                    bw.Write(0x0);
                }

                this._currentDIFATSectorPos += 4;
            }
        }
    }

    private void WritePosition(BinaryWriter bw, int sector, ref int writePos, bool isFATEntry)
    {
        int pos = (int)bw.BaseStream.Position;
        bw.Seek(writePos, SeekOrigin.Begin);
        bw.Write(sector);
        writePos += 4;
        if(isFATEntry)
        {
            this.CheckUpdateDIFAT(bw);
        }
        bw.Seek(pos, SeekOrigin.Begin);
    }
    private static void WritePosition(BinaryWriter bw, int[] sectors, ref int writePos)
    {
        int pos = (int)bw.BaseStream.Position;
        bw.Seek(writePos, SeekOrigin.Begin);
        foreach (int sector in sectors)
        {
            bw.Write(sector);
            writePos += 4;
        }
        bw.Seek(pos, SeekOrigin.Begin);
    }
    private void WriteDirs(BinaryWriter bw, List<CompoundDocumentItem> dirs)
    {
        byte[]? miniFAT = this.SetMiniStream(dirs);
        this.AllocateFAT(bw, miniFAT.Length, dirs);
        this.WriteMiniFAT(bw, miniFAT);
        foreach (CompoundDocumentItem? entity in dirs)
        {
            if (entity.ObjectType == 5 || entity.StreamSize > this._miniStreamCutoffSize)
            {
                entity.StartingSectorLocation = this.WriteStream(bw, entity.Stream);
            }
        }

        this.WriteDirStream(bw, dirs);
    }

    private int WriteDirStream(BinaryWriter bw, List<CompoundDocumentItem> dirs)
    {
        if (dirs.Count>0)
        {
            //First directory sector goes into sector 2
            bw.Seek((this._firstDirectorySectorLocation + 1) * this._sectorSize, SeekOrigin.Begin);
            for(int i=0;i<Math.Min(this._sectorSize/128,dirs.Count);i++)
            {
                dirs[i].Write(bw);
            }
        }
        else
        {
            return -1;
        }

        bw.Seek(0, SeekOrigin.End);
        int start = (int)bw.BaseStream.Position / this._sectorSize - 1;
        int pos = this._sectorSize + 4;
        this.WritePosition(bw, start, ref pos, false);
        int streamLength = 0;
        for(int i=4;i<dirs.Count;i++)
        {
            dirs[i].Write(bw);
            streamLength += 128;
        }

        WriteStreamFullSector(bw, this._sectorSize);
        this.WriteFAT(bw, start, streamLength);
        return start;

    }

    private void WriteMiniFAT(BinaryWriter bw, byte[] miniFAT)
    {
        if (miniFAT.Length >= this._sectorSize)
        {
            bw.Seek((this._firstMiniFATSectorLocation+1) * this._sectorSize, SeekOrigin.Begin);
            bw.Write(miniFAT, 0, this._sectorSize);
            bw.Seek(0, SeekOrigin.End);
            if (miniFAT.Length > this._sectorSize)
            {
                //Write next minifat sector to fat for sector 2
                int sector = ((int)bw.BaseStream.Position / this._sectorSize) - 1;
                int pos = this._sectorSize+(4*2);
                this.WritePosition(bw, sector, ref pos, false);

                //Write overflowing FAT sectors
                byte[]? b = new byte[miniFAT.Length - this._sectorSize];
                Array.Copy(miniFAT, this._sectorSize, b, 0, b.Length);
                this.WriteStream(bw, b);
            }

            this._numberofMiniFATSectors = (miniFAT.Length + 1) / this._sectorSize;
        }
    }

    private int WriteStream(BinaryWriter bw, byte[] stream)
    {
        bw.Seek(0, SeekOrigin.End);
        int start = (int)bw.BaseStream.Position / this._sectorSize-1;
        bw.Write(stream);
        WriteStreamFullSector(bw, this._sectorSize);
        this.WriteFAT(bw, start, stream.Length);
        return start;
    }
    private void WriteFAT(BinaryWriter bw, int sector, long size)
    {
        bw.Seek(this._currentFATSectorPos, SeekOrigin.Begin);
        int pos = this._sectorSize;
        while (size > pos)
        {
            bw.Write(++sector);
            pos += this._sectorSize;
            this.CheckUpdateDIFAT(bw);
        }
        bw.Write(END_OF_CHAIN);
        this.CheckUpdateDIFAT(bw);
        this._currentFATSectorPos = (int)bw.BaseStream.Position;
        bw.Seek(0, SeekOrigin.End);
    }

    private void CheckUpdateDIFAT(BinaryWriter bw)
    {
        if (bw.BaseStream.Position % this._sectorSize == 0)
        {
            if (this._currentDIFATSectorPos % this._sectorSize == 0) 
            {
                bw.Seek(512, SeekOrigin.Current);
            }
            else if (bw.BaseStream.Position == (this._sectorSize * 2))
            {
                bw.Seek(4 * this._sectorSize,SeekOrigin.Begin);    //FAT continues after initial dir och minifat sectors.
            }
            //Add to DIFAT
            int FATSector = (int)(bw.BaseStream.Position / this._sectorSize - 1);
            this.WritePosition(bw, FATSector, ref this._currentDIFATSectorPos, false);
            this._numberOfFATSectors++;
            if (this._currentDIFATSectorPos == this._sectorSize || ((this._currentDIFATSectorPos+4)  % this._sectorSize == 0 && this._currentDIFATSectorPos > this._sectorSize))
            {
                bw.Write(new byte[this._sectorSize]); //Write pre FAT sector
                if (this._currentDIFATSectorPos > this._sectorSize)                       //Write link to next DIFAT sector
                {
                    this.WritePosition(bw, FATSector+1, ref this._currentDIFATSectorPos, false);
                }
                else
                {
                    this._firstDIFATSectorLocation = FATSector+1;                    //Current sector
                }

                this._currentDIFATSectorPos = (int)bw.BaseStream.Position;
                //Fill sector
                for (int i = 0; i < this._sectorSize; i++)
                {
                    bw.Write((byte)0xFF);
                }
                bw.Seek(-(this._sectorSize * 2), SeekOrigin.Current);
            }
        }
    }

    private void AllocateFAT(BinaryWriter bw, int miniFatLength, List<CompoundDocumentItem> dirs)
    {
        /*** First calculate full size ***/
        long fullStreamSize = (long)miniFatLength - this._sectorSize; //MiniFAT starts from sector 2, by default.
        //StreamSize
        foreach (CompoundDocumentItem? entity in dirs)
        {
            if (entity.ObjectType == 5 || entity.StreamSize > this._miniStreamCutoffSize)
            {
                long rest = this._sectorSize - entity.StreamSize % this._sectorSize;
                fullStreamSize += entity.StreamSize;
                if (rest > 0 && rest < this._sectorSize)
                {
                    fullStreamSize += rest;
                }
            }
        }
        long noOfSectors = fullStreamSize / this._sectorSize;

        //Directory Size
        int dirsPerSector = this._sectorSize / 128;
        int firstFATSectorPos = this._currentFATSectorPos;
        if (dirs.Count > dirsPerSector)
        {
            int dirSectors = GetSectors(dirs.Count, dirsPerSector);
            noOfSectors += dirSectors - 1; //Four items per sector. Sector two is fixed for directories
        }

        //First calc fat no sectors and difat sectors from full size
        int numberOfFATSectors = GetSectors((int)noOfSectors, this._sectorSizeInt);       //Sector 0 is already allocated
        this._numberofDIFATSectors = this.GetDIFatSectors(numberOfFATSectors);
        noOfSectors += numberOfFATSectors + this._numberofDIFATSectors;

        //Calc fat sectors again with the added fat and di fat sectors.
        numberOfFATSectors = GetSectors((int)noOfSectors, this._sectorSizeInt) + this._numberofDIFATSectors;
        this._numberofDIFATSectors = this.GetDIFatSectors(numberOfFATSectors);

        //Allocate FAT and DIFAT Sectors
        bw.Write(new byte[(numberOfFATSectors + (this._numberofDIFATSectors > 0 ? this._numberofDIFATSectors - 1 : 0)) * this._sectorSize]);

        //Move to FAT Second sector (4).
        bw.Seek(this._currentFATSectorPos, SeekOrigin.Begin);
        int sectorPos = 1;
        for (int i = 1; i < 109; i++)     //We have 1 FAT sector to start with at sector 0
        {
            if (i < numberOfFATSectors + this._numberofDIFATSectors)
            {
                this.WriteFATItem(bw, FAT_SECTOR);
                sectorPos++;
            }
            else
            {
                this.WriteFATItem(bw, END_OF_CHAIN);
                break;
            }
        }
        if (this._numberofDIFATSectors > 0)
        {
            this._firstDIFATSectorLocation = sectorPos + 1;
        }

        for (int j = 0; j < this._numberofDIFATSectors; j++)
        {
            this.WriteFATItem(bw, DIFAT_SECTOR);
            for (int i = 0; i < this._sectorSizeInt - 1; i++)
            {
                this.WriteFATItem(bw, FAT_SECTOR);
                sectorPos++;
                if (sectorPos >= numberOfFATSectors)
                {
                    break;
                }
            }
            if (sectorPos > numberOfFATSectors)
            {
                break;
            }
        }
        bw.Seek(0, SeekOrigin.End);
    }

    private int GetDIFatSectors(int FATSectors)
    {
        if (FATSectors > 109)
        {
            return GetSectors((FATSectors - 109), this._sectorSizeInt-1);
        }
        else
        {
            return 0;
        }
    }

    private void WriteFATItem(BinaryWriter bw, int value)
    {
        bw.Write(value);
        this.CheckUpdateDIFAT(bw);
        this._currentFATSectorPos = (int)bw.BaseStream.Position;            
    }

    private static int GetSectors(int v, int size)
    {
        if(v % size==0)
        {
            return v / size;
        }
        else
        {
            return v / size + 1;
        }
    }

    private byte[] SetMiniStream(List<CompoundDocumentItem> dirs)
    {
        //Create the miniStream
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter? bwMiniFATStream = new BinaryWriter(ms);
        using MemoryStream? msforBw = RecyclableMemory.GetStream();
        BinaryWriter? bwMiniFAT = new BinaryWriter(msforBw);
        int pos = 0;
        foreach (CompoundDocumentItem? entity in dirs)
        {
            if (entity.ObjectType != 5 && entity.StreamSize > 0 && entity.StreamSize <= this._miniStreamCutoffSize)
            {
                bwMiniFATStream.Write(entity.Stream);
                WriteStreamFullSector(bwMiniFATStream, miniFATSectorSize);
                int size = this._miniSectorSize;
                entity.StartingSectorLocation = pos;
                while (entity.StreamSize > size)
                {
                    bwMiniFAT.Write(++pos);
                    size += this._miniSectorSize;
                }
                bwMiniFAT.Write(END_OF_CHAIN);
                pos++;
            }
        }
        dirs[0].StreamSize = ms.Length;
        dirs[0].Stream = ms.ToArray();
        WriteStreamFullSector(bwMiniFAT, this._sectorSize);
        return msforBw.ToArray();
    }

    private static void WriteStreamFullSector(BinaryWriter bw, int sectorSize)
    {
        long rest = sectorSize - (bw.BaseStream.Length % sectorSize);
        if (rest > 0 && rest < sectorSize)
        {
            bw.Write(new byte[rest]);
        }
    }
    private void WriteHeader(BinaryWriter bw)
    {
        bw.Seek(0, SeekOrigin.Begin);
        bw.Write(header);
        bw.Write(new byte[16]);             //ClsID all zero's
        bw.Write((short)0x3E);              //This field SHOULD be set to 0x003E if the major version field is either 0x0003 or 0x0004.                 
        bw.Write((short)0x3);               //Version 3
        bw.Write((ushort)0xFFFE);           // This field MUST be set to 0xFFFE. This field is a byte order mark for all integer fields, specifying little-endian byte order. 
        bw.Write((short)9);                 //Sector Shift
        bw.Write((short)6);                 //Mini Sector Shift
        bw.Write(new byte[6]);              //reserved
        bw.Write(0);                        //Number of Directory Sectors, unsupported i v3. Set to zero
        bw.Write(this._numberOfFATSectors);      //Number of FAT Sectors
        bw.Write(1);                        //First Directory Sector Location
        bw.Write(0);                        //Transaction Signature Number
        bw.Write(this._miniStreamCutoffSize);     //Mini Stream Cutoff Size
        bw.Write(2);                        //First Mini FAT Sector Location
        bw.Write(this._numberofMiniFATSectors);   //Number of MiniFAT sectors
        bw.Write(this._firstDIFATSectorLocation); //First DIFAT Sector Location
        bw.Write(this._numberofDIFATSectors);     //Number of DIFAT Sectors
    }

    private void CreateFATStreams(CompoundDocumentItem item, BinaryWriter bw, BinaryWriter bwMini, DocWriteInfo dwi)
    {
        if (item.ObjectType != 5)   //Root, we must have the miniStream first.
        {
            if (item.StreamSize > 0)
            {
                item.StreamSize = item.Stream.Length;
                if (item.StreamSize < this._miniStreamCutoffSize)
                {
                    item.StartingSectorLocation= WriteStream(bwMini, dwi.miniFAT, item.Stream, miniFATSectorSize);
                }
                else
                {
                    item.StartingSectorLocation = WriteStream(bw, dwi.FAT, item.Stream, FATSectorSizeV3);
                }
            }
        }
        foreach(CompoundDocumentItem? c in item.Children)
        {
            this.CreateFATStreams(c, bw, bwMini, dwi);
        }
    }

    private static int WriteStream(BinaryWriter bw, List<int> fat, byte[] stream, int FATSectorSize)
    {
        int rest = FATSectorSize - (stream.Length % FATSectorSize);
        bw.Write(stream);
        if(rest>0 && rest < FATSectorSize)
        {
            bw.Write(new byte[rest]);
        }

        int ret = fat.Count;
        AddFAT(fat, stream.Length, FATSectorSize, 0);

        return ret; //Returns the start sector
    }

    private static void AddFAT(List<int> fat, long streamSize, int sectorSize, int addPos)
    {
        int size = 0;
        while (size<streamSize)
        {
            if (size + sectorSize < streamSize)
            {
                fat.Add(fat.Count + 1);
            }
            else
            {
                fat.Add(END_OF_CHAIN);
            }
            size += sectorSize;
        }
    }

    private void LoadFromMemoryStream(MemoryStream ms)
    {
        ms.Seek(0, SeekOrigin.Begin);   //Fixes issue #60
        this.Read(new BinaryReader(ms));
    }

    public void Dispose()
    {
        this._miniSectors = null;
        this._sectors = null;            
    }
    #endregion
}