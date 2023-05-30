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
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.CompundDocument;

[DebuggerDisplay("FullName: {FullName}")]
internal class CompoundDocumentItem : IComparable<CompoundDocumentItem>
{
    public CompoundDocumentItem()
    {
        this.Children = new List<CompoundDocumentItem>();
    }

    public CompoundDocumentItem Parent { get; set; }

    public List<CompoundDocumentItem> Children { get; set; }

    public string Name { get; set; }

    public string FullName
    {
        get
        {
            string? path = this.Name;
            CompoundDocumentItem? p = this.Parent;

            while (p != null)
            {
                path = p.Name + "/" + path;
                p = p.Parent;
            }

            return path;
        }
    }

    /// <summary>
    /// 0=Red
    /// 1=Black
    /// </summary>
    public byte ColorFlag { get; set; }

    /// <summary>
    /// Type of object
    /// 0x00 - Unknown or unallocated 
    /// 0x01 - Storage Object
    /// 0x02 - Stream Object 
    /// 0x05 - Root Storage Object
    /// </summary>
    public byte ObjectType { get; set; }

    public int ChildID { get; set; }

    public Guid ClsID { get; set; }

    public int LeftSibling { get; set; }

    public int RightSibling { get; set; }

    public int StatBits { get; set; }

    public long CreationTime { get; set; }

    public long ModifiedTime { get; set; }

    public int StartingSectorLocation { get; set; }

    public long StreamSize { get; set; }

    public byte[] Stream { get; set; }

    internal bool _handled = false;

    internal void Read(BinaryReader br)
    {
        byte[]? s = br.ReadBytes(0x40);
        short sz = br.ReadInt16();

        if (sz > 0)
        {
            this.Name = Encoding.Unicode.GetString(s, 0, sz - 2);
        }

        this.ObjectType = br.ReadByte();
        this.ColorFlag = br.ReadByte();
        this.LeftSibling = br.ReadInt32();
        this.RightSibling = br.ReadInt32();
        this.ChildID = br.ReadInt32();

        //Clsid;
        this.ClsID = new Guid(br.ReadBytes(16));

        this.StatBits = br.ReadInt32();
        this.CreationTime = br.ReadInt64();
        this.ModifiedTime = br.ReadInt64();

        this.StartingSectorLocation = br.ReadInt32();
        this.StreamSize = br.ReadInt64();
    }

    internal void Write(BinaryWriter bw)
    {
        byte[]? name = Encoding.Unicode.GetBytes(this.Name);
        bw.Write(name);
        bw.Write(new byte[0x40 - name.Length]);
        bw.Write((short)(name.Length + 2));

        bw.Write(this.ObjectType);
        bw.Write(this.ColorFlag);
        bw.Write(this.LeftSibling);
        bw.Write(this.RightSibling);
        bw.Write(this.ChildID);
        bw.Write(this.ClsID.ToByteArray());
        bw.Write(this.StatBits);
        bw.Write(this.CreationTime);
        bw.Write(this.ModifiedTime);
        bw.Write(this.StartingSectorLocation);
        bw.Write(this.StreamSize);
    }

    public override string ToString()
    {
        return this.Name;
    }

    /// <summary>
    /// Compare length first, then sort by name in upper invariant
    /// </summary>
    /// <param name="other">The other item</param>
    /// <returns></returns>
    public int CompareTo(CompoundDocumentItem other)
    {
        if (this.Name.Length < other.Name.Length)
        {
            return -1;
        }
        else if (this.Name.Length > other.Name.Length)
        {
            return 1;
        }

        string? n1 = this.Name.ToUpperInvariant();
        string? n2 = other.Name.ToUpperInvariant();

        for (int i = 0; i < n1.Length; i++)
        {
            if (n1[i] < n2[i])
            {
                return -1;
            }
            else if (n1[i] > n2[i])
            {
                return 1;
            }
        }

        return 0;
    }
}