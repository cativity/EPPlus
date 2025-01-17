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
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml.Utils;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.Constants;
using System.Collections.Generic;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.VBA;

/// <summary>
/// Represents the VBA project part of the package
/// </summary>
public class ExcelVbaProject
{
    const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    internal const string PartUri = @"/xl/vbaProject.bin";

    internal ExcelVbaProject(ExcelWorkbook wb)
    {
        this._wb = wb;
        this._pck = this._wb._package.ZipPackage;
        this.References = new ExcelVbaReferenceCollection();
        this.Modules = new ExcelVbaModuleCollection(this);
        ZipPackageRelationship? rel = this._wb.Part.GetRelationshipsByType(schemaRelVba).FirstOrDefault();

        if (rel != null)
        {
            this.Uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            this.Part = this._pck.GetPart(this.Uri);
            this.GetProject();
        }
        else
        {
            this.Lcid = 0;
            this.Part = null;
        }
    }

    internal ExcelWorkbook _wb;
    internal ZipPackage _pck;

    #region Dir Stream Properties

    /// <summary>
    /// System kind. Default Win32.
    /// </summary>
    public eSyskind SystemKind { get; set; }

    /// <summary>
    /// The compatible version for the VBA project. If null, this record is not written.
    /// </summary>
    public uint? CompatVersion { get; set; }

    /// <summary>
    /// Name of the project
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// A description of the project
    /// </summary>
    public string Description { get; set; }

    /// <summary>
    /// A helpfile
    /// </summary>
    public string HelpFile1 { get; set; }

    /// <summary>
    /// Secondary helpfile
    /// </summary>
    public string HelpFile2 { get; set; }

    /// <summary>
    /// Context if refering the helpfile
    /// </summary>
    public int HelpContextID { get; set; }

    /// <summary>
    /// Conditional compilation constants
    /// </summary>
    public string Constants { get; set; }

    /// <summary>
    /// Codepage for encoding. Default is current regional setting.
    /// </summary>
    public int CodePage { get; internal set; }

    internal int LibFlags { get; set; }

    internal int MajorVersion { get; set; }

    internal int MinorVersion { get; set; }

    internal int Lcid { get; set; }

    internal int LcidInvoke { get; set; }

    internal string ProjectID { get; set; }

    internal string ProjectStreamText { get; set; }

    /// <summary>
    /// Project references
    /// </summary>        
    public ExcelVbaReferenceCollection References { get; set; }

    /// <summary>
    /// Code Modules (Modules, classes, designer code)
    /// </summary>
    public ExcelVbaModuleCollection Modules { get; set; }

    internal List<string> _HostExtenders = new List<string>();
    ExcelVbaSignature _signature;

    /// <summary>
    /// The digital signature
    /// </summary>
    public ExcelVbaSignature Signature => this._signature ??= new ExcelVbaSignature(this.Part);

    ExcelVbaProtection _protection;

    /// <summary>
    /// VBA protection 
    /// </summary>
    public ExcelVbaProtection Protection => this._protection ??= new ExcelVbaProtection(this);

    #endregion

    #region Read Project

    private void GetProject()
    {
        Stream? stream = this.Part.GetStream();
        byte[] vba = new byte[stream.Length];
        _ = stream.Read(vba, 0, (int)stream.Length);
        this.Document = new CompoundDocument(vba);

        this.ReadDirStream();
        this.ProjectStreamText = Encoding.GetEncoding(this.CodePage).GetString(this.Document.Storage.DataStreams["PROJECT"]);
        this.ReadModules();
        this.ReadProjectProperties();
    }

    private void ReadModules()
    {
        foreach (ExcelVBAModule? modul in this.Modules)
        {
            byte[]? stream = this.Document.Storage.SubStorage["VBA"].DataStreams[modul.streamName];
            byte[]? byCode = VBACompression.DecompressPart(stream, (int)modul.ModuleOffset);
            string code = Encoding.GetEncoding(this.CodePage).GetString(byCode);
            int pos = 0;

            while (pos + 9 < code.Length && code.Substring(pos, 9) == "Attribute")
            {
                int linePos = code.IndexOf("\r\n", pos, StringComparison.OrdinalIgnoreCase);
                int crlfSize;

                if (linePos < 0)
                {
                    linePos = code.IndexOf("\n", pos, StringComparison.OrdinalIgnoreCase);
                    crlfSize = 1;
                }
                else
                {
                    crlfSize = 2;
                }

                string[] lineSplit;

                if (linePos > 0)
                {
                    lineSplit = code.Substring(pos + 9, linePos - pos - 9).Split('=');
                }
                else
                {
                    lineSplit = code.Substring(pos + 9).Split(new char[] { '=' }, 1);
                }

                if (lineSplit.Length > 1)
                {
                    lineSplit[1] = lineSplit[1].Trim();

                    ExcelVbaModuleAttribute? attr = new ExcelVbaModuleAttribute()
                    {
                        Name = lineSplit[0].Trim(),
                        DataType =
                            lineSplit[1].StartsWith("\"", StringComparison.OrdinalIgnoreCase) ? eAttributeDataType.String : eAttributeDataType.NonString,
                        Value = lineSplit[1].StartsWith("\"", StringComparison.OrdinalIgnoreCase)
                                    ? lineSplit[1].Substring(1, lineSplit[1].Length - 2)
                                    : lineSplit[1]
                    };

                    modul.Attributes._list.Add(attr);
                }

                pos = linePos + crlfSize;
            }

            modul.Code = code.Substring(pos);
        }
    }

    private void ReadProjectProperties()
    {
        this._protection = new ExcelVbaProtection(this);
        string prevPackage = "";
        string[]? lines = Regex.Split(this.ProjectStreamText, "\r\n");

        bool isHostExtender = false,
             isWorkspace = false;

        foreach (string line in lines)
        {
            if (line.StartsWith("[", StringComparison.OrdinalIgnoreCase))
            {
                switch (line.Trim())
                {
                    case "[Host Extender Info]":
                        isHostExtender = true;
                        this._HostExtenders.Clear();

                        break;

                    case "[Workspace]":
                        isWorkspace = true;

                        break;
                }
            }
            else if (isWorkspace)
            {
                //We ignore workspaces for now and set all windows with coordinates 0,0.
            }
            else if (isHostExtender)
            {
                if (string.IsNullOrEmpty(line) == false)
                {
                    this._HostExtenders.Add(line);
                }
            }
            else
            {
                string[]? split = line.Split('=');

                if (split.Length > 1 && split[1].Length > 1 && split[1].StartsWith("\"", StringComparison.OrdinalIgnoreCase)) //Remove any double qouates
                {
                    split[1] = split[1].Substring(1, split[1].Length - 2);
                }

                switch (split[0])
                {
                    case "ID":
                        this.ProjectID = split[1];

                        break;

                    case "Document":
                        string mn = split[1].Substring(0, split[1].IndexOf("/&H", StringComparison.OrdinalIgnoreCase));

                        if (this.Modules.Exists(mn))
                        {
                            this.Modules[mn].Type = eModuleType.Document;
                        }

                        break;

                    case "Package":
                        prevPackage = split[1];

                        break;

                    case "BaseClass":
                        if (this.Modules.Exists(split[1]))
                        {
                            this.Modules[split[1]].Type = eModuleType.Designer;
                            this.Modules[split[1]].ClassID = prevPackage;
                        }

                        break;

                    case "Module":
                        if (this.Modules.Exists(split[1]))
                        {
                            this.Modules[split[1]].Type = eModuleType.Module;
                        }

                        break;

                    case "Class":
                        if (this.Modules.Exists(split[1]))
                        {
                            this.Modules[split[1]].Type = eModuleType.Class;
                        }

                        break;

                    case "HelpFile":
                    case "Name":
                    case "HelpContextID":
                    case "Description":
                    case "VersionCompatible32":
                        break;

                    //393222000"
                    case "CMG":
                        byte[] cmg = Decrypt(split[1]);
                        this._protection.UserProtected = (cmg[0] & 1) != 0;
                        this._protection.HostProtected = (cmg[0] & 2) != 0;
                        this._protection.VbeProtected = (cmg[0] & 4) != 0;

                        break;

                    case "DPB":
                        byte[] dpb = Decrypt(split[1]);

                        if (dpb.Length >= 28)
                        {
                            byte[]? flags = new byte[3];
                            Array.Copy(dpb, 1, flags, 0, 3);
                            byte[]? keyNoNulls = new byte[4];
                            this._protection.PasswordKey = new byte[4];
                            Array.Copy(dpb, 4, keyNoNulls, 0, 4);
                            byte[]? hashNoNulls = new byte[20];
                            this._protection.PasswordHash = new byte[20];
                            Array.Copy(dpb, 8, hashNoNulls, 0, 20);

                            //Handle 0x00 bitwise 2.4.4.3 
                            for (int i = 0; i < 24; i++)
                            {
                                int bit = 128 >> (int)(i % 8);

                                if (i < 4)
                                {
                                    if ((int)(flags[0] & bit) == 0)
                                    {
                                        this._protection.PasswordKey[i] = 0;
                                    }
                                    else
                                    {
                                        this._protection.PasswordKey[i] = keyNoNulls[i];
                                    }
                                }
                                else
                                {
                                    int flagIndex = (i - (i % 8)) / 8;

                                    if ((int)(flags[flagIndex] & bit) == 0)
                                    {
                                        this._protection.PasswordHash[i - 4] = 0;
                                    }
                                    else
                                    {
                                        this._protection.PasswordHash[i - 4] = hashNoNulls[i - 4];
                                    }
                                }
                            }
                        }

                        break;

                    case "GC":
                        this._protection.VisibilityState = Decrypt(split[1])[0] == 0xFF;

                        break;
                }
            }
        }
    }

    /// <summary>
    /// 2.4.3.3 Decryption
    /// </summary>
    /// <param name="value">Byte hex string</param>
    /// <returns>The decrypted value</returns>
    private static byte[] Decrypt(string value)
    {
        byte[] enc = GetByte(value);
        byte[] dec = new byte[value.Length - 1];
        byte seed = enc[0];
        dec[0] = (byte)(enc[1] ^ seed);
        dec[1] = (byte)(enc[2] ^ seed);

        for (int i = 2; i < enc.Length - 1; i++)
        {
            dec[i] = (byte)(enc[i + 1] ^ (enc[i - 1] + dec[i - 1]));
        }

        byte ignoredLength = (byte)((seed & 6) / 2);
        int datalength = BitConverter.ToInt32(dec, ignoredLength + 2);
        byte[]? data = new byte[datalength];
        Array.Copy(dec, 6 + ignoredLength, data, 0, datalength);

        return data;
    }

    /// <summary>
    /// 2.4.3.2 Encryption
    /// </summary>
    /// <param name="value"></param>
    /// <returns>Byte hex string</returns>
    private string Encrypt(byte[] value)
    {
        byte[] seed = new byte[1];
        RandomNumberGenerator? rn = RandomNumberGenerator.Create();
        rn.GetBytes(seed);

        byte[] array;
        byte pb;
        byte[] enc = new byte[value.Length + 10];

        using (MemoryStream? ms = RecyclableMemory.GetStream())
        {
            BinaryWriter br = new BinaryWriter(ms);
            enc[0] = seed[0];
            enc[1] = (byte)(2 ^ seed[0]);

            byte projKey = 0;

            foreach (char c in this.ProjectID)
            {
                projKey += (byte)c;
            }

            enc[2] = (byte)(projKey ^ seed[0]);
            int ignoredLength = (seed[0] & 6) / 2;

            for (int i = 0; i < ignoredLength; i++)
            {
                br.Write(seed[0]);
            }

            br.Write(value.Length);
            br.Write(value);
            array = ms.ToArray();
            pb = projKey;
        }

        int pos = 3;

        foreach (byte b in array)
        {
            enc[pos] = (byte)(b ^ (enc[pos - 2] + pb));
            pos++;
            pb = b;
        }

        return GetString(enc, pos - 1);
    }

    private static string GetString(byte[] value, int max)
    {
        string ret = "";

        for (int i = 0; i <= max; i++)
        {
            if (value[i] < 16)
            {
                ret += "0" + value[i].ToString("x");
            }
            else
            {
                ret += value[i].ToString("x");
            }
        }

        return ret.ToUpperInvariant();
    }

    private static byte[] GetByte(string value)
    {
        byte[] ret = new byte[value.Length / 2];

        for (int i = 0; i < ret.Length; i++)
        {
            ret[i] = byte.Parse(value.Substring(i * 2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
        }

        return ret;
    }

    private void ReadDirStream()
    {
        byte[] dir = VBACompression.DecompressPart(this.Document.Storage.SubStorage["VBA"].DataStreams["dir"]);
        using MemoryStream? ms = RecyclableMemory.GetStream(dir);
        BinaryReader br = new BinaryReader(ms);
        ExcelVbaReference currentRef = null;
        string referenceName = "";
        ExcelVBAModule currentModule = null;
        bool terminate = false;

        while (ms.Position < ms.Length && terminate == false)
        {
            ushort id = br.ReadUInt16();
            uint size = br.ReadUInt32();

            switch (id)
            {
                case 0x01:
                    this.SystemKind = (eSyskind)br.ReadUInt32();

                    break;

                case 0x02:
                    this.Lcid = (int)br.ReadUInt32();

                    break;

                case 0x03:
                    this.CodePage = (int)br.ReadUInt16();

                    break;

                case 0x04:
                    this.Name = this.GetString(br, size);

                    break;

                case 0x05:
                    this.Description = this.GetStringAndUnicodeString(br, size);

                    break;

                case 0x06:
                    this.HelpFile1 = this.GetString(br, size);

                    break;

                case 0x3D:
                    this.HelpFile2 = this.GetString(br, size);

                    break;

                case 0x07:
                    this.HelpContextID = (int)br.ReadUInt32();

                    break;

                case 0x08:
                    this.LibFlags = (int)br.ReadUInt32();

                    break;

                case 0x09:
                    this.MajorVersion = (int)br.ReadUInt32();
                    this.MinorVersion = (int)br.ReadUInt16();

                    break;

                case 0x0C:
                    this.Constants = this.GetStringAndUnicodeString(br, size);

                    break;

                case 0x0D:
                    uint sizeLibID = br.ReadUInt32();
                    ExcelVbaReference? regRef = new ExcelVbaReference();
                    regRef.Name = referenceName;
                    regRef.ReferenceRecordID = id;
                    regRef.Libid = this.GetString(br, sizeLibID);
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt16();
                    this.References.Add(regRef);

                    break;

                case 0x0E:
                    ExcelVbaReferenceProject? projRef = new ExcelVbaReferenceProject();
                    projRef.ReferenceRecordID = id;
                    projRef.Name = referenceName;
                    sizeLibID = br.ReadUInt32();
                    projRef.Libid = this.GetString(br, sizeLibID);
                    sizeLibID = br.ReadUInt32();
                    projRef.LibIdRelative = this.GetString(br, sizeLibID);
                    projRef.MajorVersion = br.ReadUInt32();
                    projRef.MinorVersion = br.ReadUInt16();
                    this.References.Add(projRef);

                    break;

                case 0x0F:
                    _ = br.ReadUInt16();

                    break;

                case 0x13:
                    _ = br.ReadUInt16();

                    break;

                case 0x14:
                    this.LcidInvoke = (int)br.ReadUInt32();

                    break;

                case 0x16:
                    referenceName = this.GetStringAndUnicodeString(br, size);

                    break;

                case 0x19:
                    currentModule = new ExcelVBAModule();
                    currentModule.Name = this.GetString(br, size);
                    this.Modules.Add(currentModule);

                    break;

                case 0x47:
                    currentModule.NameUnicode = GetString(br, size, Encoding.Unicode);

                    break;

                case 0x1A:
                    currentModule.streamName = this.GetStringAndUnicodeString(br, size);

                    break;

                case 0x1C:
                    currentModule.Description = this.GetStringAndUnicodeString(br, size);

                    break;

                case 0x1E:
                    currentModule.HelpContext = (int)br.ReadUInt32();

                    break;

                case 0x21:
                case 0x22:
                    break;

                case 0x2B: //Modul Terminator
                    break;

                case 0x2C:
                    currentModule.Cookie = br.ReadUInt16();

                    break;

                case 0x31:
                    currentModule.ModuleOffset = br.ReadUInt32();

                    break;

                case 0x10:
                    terminate = true;

                    break;

                case 0x30:
                    ExcelVbaReferenceControl? extRef = (ExcelVbaReferenceControl)currentRef;
                    uint sizeExt = br.ReadUInt32();
                    extRef.LibIdExtended = this.GetString(br, sizeExt);

                    _ = br.ReadUInt32();
                    _ = br.ReadUInt16();
                    extRef.OriginalTypeLib = new Guid(br.ReadBytes(16));
                    extRef.Cookie = br.ReadUInt32();

                    break;

                case 0x33:
                    currentRef = new ExcelVbaReferenceControl();
                    currentRef.ReferenceRecordID = id;
                    currentRef.Name = referenceName;
                    currentRef.Libid = this.GetString(br, size);
                    this.References.Add(currentRef);

                    break;

                case 0x2F:
                    ExcelVbaReferenceControl? contrRef = (ExcelVbaReferenceControl)currentRef;
                    contrRef.SecondaryReferenceRecordID = id;

                    uint sizeTwiddled = br.ReadUInt32();
                    contrRef.LibIdTwiddled = this.GetString(br, sizeTwiddled);
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt16();

                    break;

                case 0x25:
                    currentModule.ReadOnly = true;

                    break;

                case 0x28:
                    currentModule.Private = true;

                    break;

                case 0x4a:
                    this.CompatVersion = br.ReadUInt32();

                    break;

                default:
                    _ = br.ReadBytes((int)size);

                    break;
            }
        }
    }

    #endregion

    #region Save Project

    internal void Save()
    {
        if (this.Validate())
        {
            CompoundDocument doc = new CompoundDocument();
            doc.Storage = new CompoundDocument.StoragePart();
            CompoundDocument.StoragePart? store = new CompoundDocument.StoragePart();
            doc.Storage.SubStorage.Add("VBA", store);

            store.DataStreams.Add("_VBA_PROJECT", CreateVBAProjectStream());
            store.DataStreams.Add("dir", this.CreateDirStream());

            foreach (ExcelVBAModule? module in this.Modules)
            {
                module.ModuleOffset = 0;

                store.DataStreams.Add(module.Name,
                                      VBACompression.CompressPart(Encoding.GetEncoding(this.CodePage)
                                                                          .GetBytes(module.Attributes.GetAttributeText() + module.Code)));
            }

            //Copy streams from the template, if used.
            if (this.Document != null)
            {
                foreach (KeyValuePair<string, CompoundDocument.StoragePart> ss in this.Document.Storage.SubStorage)
                {
                    if (ss.Key != "VBA")
                    {
                        doc.Storage.SubStorage.Add(ss.Key, ss.Value);
                    }
                }

                foreach (KeyValuePair<string, byte[]> s in this.Document.Storage.DataStreams)
                {
                    if (s.Key != "dir" && s.Key != "PROJECT" && s.Key != "PROJECTwm")
                    {
                        doc.Storage.DataStreams.Add(s.Key, s.Value);
                    }
                }
            }

            doc.Storage.DataStreams.Add("PROJECT", this.CreateProjectStream());
            doc.Storage.DataStreams.Add("PROJECTwm", this.CreateProjectwmStream());

            if (this.Part == null)
            {
                this.Uri = new Uri(PartUri, UriKind.Relative);
                this.Part = this._pck.CreatePart(this.Uri, ContentTypes.contentTypeVBA);
                _ = this._wb.Part.CreateRelationship(this.Uri, TargetMode.Internal, schemaRelVba);
            }

            Stream? st = this.Part.GetStream(FileMode.Create);
            doc.Save((MemoryStream)st);

            this.Document = doc;
            st.Flush();

            //Save the digital signture
            this.Signature.Save(this);
        }
    }

    private bool Validate()
    {
        this.Description ??= "";
        this.HelpFile1 ??= "";
        this.HelpFile2 ??= "";
        this.Constants ??= "";

        return true;
    }

    /// <summary>
    /// MS-OVBA 2.3.4.1
    /// </summary>
    /// <returns></returns>
    private static byte[] CreateVBAProjectStream()
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter bw = new BinaryWriter(ms);
        bw.Write((ushort)0x61CC); //Reserved1
        bw.Write((ushort)0xFFFF); //Version
        bw.Write((byte)0x0); //Reserved3
        bw.Write((ushort)0x0); //Reserved4

        return ms.ToArray();
    }

    /// <summary>
    /// MS-OVBA 2.3.4.1
    /// </summary>
    /// <returns></returns>
    private byte[] CreateDirStream()
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter bw = new BinaryWriter(ms);

        /****** PROJECTINFORMATION Record ******/
        bw.Write((ushort)1); //ID
        bw.Write((uint)4); //Size
        bw.Write((uint)this.SystemKind); //SysKind

        if (this.CompatVersion.HasValue)
        {
            bw.Write((ushort)0x4a); //ID
            bw.Write((uint)4); //Size
            bw.Write((uint)this.CompatVersion.Value); //compatversion
        }

        bw.Write((ushort)2); //ID
        bw.Write((uint)4); //Size
        bw.Write((uint)this.Lcid); //Lcid

        bw.Write((ushort)0x14); //ID
        bw.Write((uint)4); //Size
        bw.Write((uint)this.LcidInvoke); //Lcid Invoke

        bw.Write((ushort)3); //ID
        bw.Write((uint)2); //Size
        bw.Write((ushort)this.CodePage); //Codepage

        //ProjectName
        bw.Write((ushort)4); //ID
        byte[]? nameBytes = Encoding.GetEncoding(this.CodePage).GetBytes(this.Name);
        bw.Write((uint)nameBytes.Length); //Size
        bw.Write(nameBytes); //Project Name

        //Description
        bw.Write((ushort)5); //ID
        byte[]? descriptionBytes = Encoding.GetEncoding(this.CodePage).GetBytes(this.Description);
        bw.Write((uint)descriptionBytes.Length); //Size
        bw.Write(descriptionBytes); //Project Name
        bw.Write((ushort)0x40); //ID
        byte[]? descriptionUnicodeBytes = Encoding.Unicode.GetBytes(this.Description);
        bw.Write((uint)descriptionUnicodeBytes.Length); //Size
        bw.Write(descriptionUnicodeBytes); //Project Description

        //Helpfiles
        bw.Write((ushort)6); //ID
        byte[]? helpFile1Bytes = Encoding.GetEncoding(this.CodePage).GetBytes(this.HelpFile1);
        bw.Write((uint)helpFile1Bytes.Length); //Size
        bw.Write(helpFile1Bytes); //HelpFile1            
        bw.Write((ushort)0x3D); //ID
        byte[]? helpFile2Bytes = Encoding.GetEncoding(this.CodePage).GetBytes(this.HelpFile2);
        bw.Write((uint)helpFile2Bytes.Length); //Size
        bw.Write(helpFile2Bytes); //HelpFile2

        //Help context id
        bw.Write((ushort)7); //ID
        bw.Write((uint)4); //Size
        bw.Write((uint)this.HelpContextID); //Help context id

        //Libflags
        bw.Write((ushort)8); //ID
        bw.Write((uint)4); //Size
        bw.Write((uint)0); //Help context id

        //Vba Version
        bw.Write((ushort)9); //ID
        bw.Write((uint)4); //Reserved
        bw.Write((uint)this.MajorVersion); //Reserved
        bw.Write((ushort)this.MinorVersion); //Help context id

        //Constants
        bw.Write((ushort)0x0C); //ID

        byte[]? constantsBytes = Encoding.GetEncoding(this.CodePage).GetBytes(this.Constants);
        bw.Write((uint)constantsBytes.Length); //Size
        bw.Write(constantsBytes);

        byte[]? constantsUnicodeBytes = Encoding.Unicode.GetBytes(this.Constants);
        bw.Write((ushort)0x3C); //ID
        bw.Write((uint)constantsUnicodeBytes.Length); //Size
        bw.Write(constantsUnicodeBytes); //

        /****** PROJECTREFERENCES Record ******/
        foreach (ExcelVbaReference? reference in this.References)
        {
            this.WriteNameReference(bw, reference);

            if (reference.SecondaryReferenceRecordID == 0x2F)
            {
                this.WriteControlReference(bw, reference);
            }
            else if (reference.ReferenceRecordID == 0x33)
            {
                this.WriteOrginalReference(bw, reference);
            }
            else if (reference.ReferenceRecordID == 0x0D)
            {
                this.WriteRegisteredReference(bw, reference);
            }
            else if (reference.ReferenceRecordID == 0x0E)
            {
                this.WriteProjectReference(bw, reference);
            }
        }

        bw.Write((ushort)0x0F);
        bw.Write((uint)0x02);
        bw.Write((ushort)this.Modules.Count);
        bw.Write((ushort)0x13);
        bw.Write((uint)0x02);
        bw.Write((ushort)0xFFFF);

        foreach (ExcelVBAModule? module in this.Modules)
        {
            this.WriteModuleRecord(bw, module);
        }

        bw.Write((ushort)0x10); //Terminator
        bw.Write((uint)0);

        return VBACompression.CompressPart(ms.ToArray());
    }

    private void WriteModuleRecord(BinaryWriter bw, ExcelVBAModule module)
    {
        bw.Write((ushort)0x19);
        byte[]? nameBytes = Encoding.GetEncoding(this.CodePage).GetBytes(module.Name);
        bw.Write((uint)nameBytes.Length);
        bw.Write(nameBytes); //Name

        bw.Write((ushort)0x47);
        byte[]? nameUnicodeBytes = Encoding.Unicode.GetBytes(module.Name);
        bw.Write((uint)nameUnicodeBytes.Length);
        bw.Write(nameUnicodeBytes); //Name

        bw.Write((ushort)0x1A);
        bw.Write((uint)nameBytes.Length);
        bw.Write(nameBytes); //Stream Name  

        bw.Write((ushort)0x32);
        bw.Write((uint)nameUnicodeBytes.Length);
        bw.Write(nameUnicodeBytes); //Stream Name

        module.Description ??= "";
        bw.Write((ushort)0x1C);
        byte[]? descriptionBytes = Encoding.GetEncoding(this.CodePage).GetBytes(module.Description);
        bw.Write((uint)descriptionBytes.Length);
        bw.Write(descriptionBytes); //Description

        bw.Write((ushort)0x48);
        byte[]? descriptionUnicodeBytes = Encoding.Unicode.GetBytes(module.Description);
        bw.Write((uint)descriptionUnicodeBytes.Length);
        bw.Write(descriptionUnicodeBytes); //Description

        bw.Write((ushort)0x31);
        bw.Write((uint)4);
        bw.Write((uint)0); //Module Stream Offset (No PerformanceCache)

        bw.Write((ushort)0x1E);
        bw.Write((uint)4);
        bw.Write((uint)module.HelpContext); //Help context ID

        bw.Write((ushort)0x2C);
        bw.Write((uint)2);
        bw.Write((ushort)0xFFFF); //Help context ID

        bw.Write((ushort)(module.Type == eModuleType.Module ? 0x21 : 0x22));
        bw.Write((uint)0);

        if (module.ReadOnly)
        {
            bw.Write((ushort)0x25);
            bw.Write((uint)0); //Readonly
        }

        if (module.Private)
        {
            bw.Write((ushort)0x28);
            bw.Write((uint)0); //Private
        }

        bw.Write((ushort)0x2B); //Terminator
        bw.Write((uint)0);
    }

    private void WriteNameReference(BinaryWriter bw, ExcelVbaReference reference)
    {
        //Name record
        bw.Write((ushort)0x16); //ID
        byte[]? nameBytes = Encoding.GetEncoding(this.CodePage).GetBytes(reference.Name);
        bw.Write((uint)nameBytes.Length); //Size
        bw.Write(nameBytes); //HelpFile1

        bw.Write((ushort)0x3E); //ID

        byte[]? nameUnicodeBytes = Encoding.Unicode.GetBytes(reference.Name);
        bw.Write((uint)nameUnicodeBytes.Length); //Size
        bw.Write(nameUnicodeBytes); //HelpFile2
    }

    private void WriteControlReference(BinaryWriter bw, ExcelVbaReference reference)
    {
        this.WriteOrginalReference(bw, reference);

        bw.Write((ushort)0x2F);
        ExcelVbaReferenceControl? controlRef = (ExcelVbaReferenceControl)reference;

        byte[]? libIdTwiddledBytes = Encoding.GetEncoding(this.CodePage).GetBytes(controlRef.LibIdTwiddled);
        bw.Write((uint)(4 + libIdTwiddledBytes.Length + 4 + 2)); // Size of SizeOfLibidTwiddled, LibidTwiddled, Reserved1, and Reserved2.
        bw.Write((uint)libIdTwiddledBytes.Length); //Size            
        bw.Write(libIdTwiddledBytes); //LibID

        bw.Write((uint)0); //Reserved1
        bw.Write((ushort)0); //Reserved2
        this.WriteNameReference(bw, reference); //Name record again
        bw.Write((ushort)0x30); //Reserved3

        byte[]? libIdExternalBytes = Encoding.GetEncoding(this.CodePage).GetBytes(controlRef.LibIdExtended);

        bw.Write((uint)(4
                        + libIdExternalBytes.Length
                        + 4
                        + 2
                        + 16
                        + 4)); //Size of SizeOfLibidExtended, LibidExtended, Reserved4, Reserved5, OriginalTypeLib, and Cookie

        bw.Write((uint)libIdExternalBytes.Length); //Size            
        bw.Write(libIdExternalBytes); //LibID
        bw.Write((uint)0); //Reserved4
        bw.Write((ushort)0); //Reserved5
        bw.Write(controlRef.OriginalTypeLib.ToByteArray());
        bw.Write((uint)controlRef.Cookie); //Cookie
    }

    private void WriteOrginalReference(BinaryWriter bw, ExcelVbaReference reference)
    {
        bw.Write((ushort)0x33);
        byte[]? libIdBytes = Encoding.GetEncoding(this.CodePage).GetBytes(reference.Libid);
        bw.Write((uint)libIdBytes.Length);
        bw.Write(libIdBytes); //LibID
    }

    private void WriteProjectReference(BinaryWriter bw, ExcelVbaReference reference)
    {
        bw.Write((ushort)0x0E);
        ExcelVbaReferenceProject? projRef = (ExcelVbaReferenceProject)reference;
        byte[]? libIdBytes = Encoding.GetEncoding(this.CodePage).GetBytes(projRef.Libid);
        byte[]? libIdRelativeBytes = Encoding.GetEncoding(this.CodePage).GetBytes(projRef.LibIdRelative);
        bw.Write((uint)(4 + libIdBytes.Length + 4 + libIdRelativeBytes.Length + 4 + 2));
        bw.Write((uint)libIdBytes.Length);
        bw.Write(libIdBytes); //LibAbsolute
        bw.Write((uint)libIdRelativeBytes.Length);
        bw.Write(libIdRelativeBytes); //LibIdRelative
        bw.Write(projRef.MajorVersion);
        bw.Write(projRef.MinorVersion);
    }

    private void WriteRegisteredReference(BinaryWriter bw, ExcelVbaReference reference)
    {
        bw.Write((ushort)0x0D);
        byte[]? libIdBytes = Encoding.GetEncoding(this.CodePage).GetBytes(reference.Libid);
        bw.Write((uint)(4 + libIdBytes.Length + 4 + 2));
        bw.Write((uint)libIdBytes.Length);
        bw.Write(libIdBytes); //LibID            
        bw.Write((uint)0); //Reserved1
        bw.Write((ushort)0); //Reserved2
    }

    private byte[] CreateProjectwmStream()
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter bw = new BinaryWriter(ms);

        foreach (ExcelVBAModule? module in this.Modules)
        {
            bw.Write(Encoding.GetEncoding(this.CodePage).GetBytes(module.Name)); //Name
            bw.Write((byte)0); //Null
            bw.Write(Encoding.Unicode.GetBytes(module.Name)); //Name
            bw.Write((ushort)0); //Null
        }

        bw.Write((ushort)0); //Null

        return ms.ToArray();
    }

    private byte[] CreateProjectStream()
    {
        StringBuilder sb = new StringBuilder();
        _ = sb.AppendFormat("ID=\"{0}\"\r\n", this.ProjectID);

        foreach (ExcelVBAModule? module in this.Modules)
        {
            if (module.Type == eModuleType.Document)
            {
                _ = sb.AppendFormat("Document={0}/&H00000000\r\n", module.Name);
            }
            else if (module.Type == eModuleType.Module)
            {
                _ = sb.AppendFormat("Module={0}\r\n", module.Name);
            }
            else if (module.Type == eModuleType.Class)
            {
                _ = sb.AppendFormat("Class={0}\r\n", module.Name);
            }
            else
            {
                //Designer
                if (string.IsNullOrEmpty(module.ClassID) == false)
                {
                    _ = sb.AppendFormat("Package={0}\r\n", module.ClassID);
                }

                _ = sb.AppendFormat("BaseClass={0}\r\n", module.Name);
            }
        }

        if (this.HelpFile1 != "")
        {
            _ = sb.AppendFormat("HelpFile={0}\r\n", this.HelpFile1);
        }

        _ = sb.AppendFormat("Name=\"{0}\"\r\n", this.Name);
        _ = sb.AppendFormat("HelpContextID={0}\r\n", this.HelpContextID);

        if (!string.IsNullOrEmpty(this.Description))
        {
            _ = sb.AppendFormat("Description=\"{0}\"\r\n", this.Description);
        }

        _ = sb.AppendFormat("VersionCompatible32=\"393222000\"\r\n");

        _ = sb.AppendFormat("CMG=\"{0}\"\r\n", this.WriteProtectionStat());
        _ = sb.AppendFormat("DPB=\"{0}\"\r\n", this.WritePassword());
        _ = sb.AppendFormat("GC=\"{0}\"\r\n\r\n", this.WriteVisibilityState());

        _ = sb.Append("[Host Extender Info]\r\n");

        if (this._HostExtenders.Count == 0)
        {
            _ = sb.Append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
        }
        else
        {
            foreach (string? line in this._HostExtenders)
            {
                _ = sb.Append($"{line}\r\n");
            }
        }

        _ = sb.Append("\r\n");
        _ = sb.Append("[Workspace]\r\n");

        foreach (ExcelVBAModule? module in this.Modules)
        {
            _ = sb.AppendFormat("{0}=0, 0, 0, 0, C \r\n", module.Name);
        }

        string s = sb.ToString();

        return Encoding.GetEncoding(this.CodePage).GetBytes(s);
    }

    private string WriteProtectionStat()
    {
        int stat = (this._protection.UserProtected ? 1 : 0) | (this._protection.HostProtected ? 2 : 0) | (this._protection.VbeProtected ? 4 : 0);

        return this.Encrypt(BitConverter.GetBytes(stat));
    }

    private string WritePassword()
    {
        byte[] nullBits = new byte[3];
        byte[] nullKey = new byte[4];
        byte[] nullHash = new byte[20];

        if (this.Protection.PasswordKey == null)
        {
            return this.Encrypt(new byte[] { 0 });
        }
        else
        {
            Array.Copy(this.Protection.PasswordKey, nullKey, 4);
            Array.Copy(this.Protection.PasswordHash, nullHash, 20);

            //Set Null bits
            for (int i = 0; i < 24; i++)
            {
                byte bit = (byte)(128 >> (int)(i % 8));

                if (i < 4)
                {
                    if (nullKey[i] == 0)
                    {
                        nullKey[i] = 1;
                    }
                    else
                    {
                        nullBits[0] |= bit;
                    }
                }
                else
                {
                    if (nullHash[i - 4] == 0)
                    {
                        nullHash[i - 4] = 1;
                    }
                    else
                    {
                        int byteIndex = (i - (i % 8)) / 8;
                        nullBits[byteIndex] |= bit;
                    }
                }
            }

            //Write the Password Hash Data Structure (2.4.4.1)
            using MemoryStream? ms = RecyclableMemory.GetStream();
            BinaryWriter bw = new BinaryWriter(ms);
            bw.Write((byte)0xFF);
            bw.Write(nullBits);
            bw.Write(nullKey);
            bw.Write(nullHash);
            bw.Write((byte)0);

            return this.Encrypt(ms.ToArray());
        }
    }

    private string WriteVisibilityState() => this.Encrypt(new byte[] { (byte)(this.Protection.VisibilityState ? 0xFF : 0) });

    #endregion

    private string GetString(BinaryReader br, uint size) => GetString(br, size, Encoding.GetEncoding(this.CodePage));

    private static string GetString(BinaryReader br, uint size, Encoding enc)
    {
        if (size > 0)
        {
            byte[] byteTemp = br.ReadBytes((int)size);

            return enc.GetString(byteTemp);
        }
        else
        {
            return "";
        }
    }

    private string GetStringAndUnicodeString(BinaryReader br, uint size)
    {
        string s = this.GetString(br, size);
        _ = br.ReadUInt16();
        uint sizeUC = br.ReadUInt32();
        string sUC = GetString(br, sizeUC, Encoding.Unicode);

        return sUC.Length == 0 ? s : sUC;
    }

    internal CompoundDocument Document { get; set; }

    internal ZipPackagePart Part { get; set; }

    internal Uri Uri { get; private set; }

    /// <summary>
    /// Create a new VBA Project
    /// </summary>
    internal void Create()
    {
        if (this.Lcid > 0)
        {
            throw new InvalidOperationException("Package already contains a VBAProject");
        }

        this.ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
        this.Name = "VBAProject";
        this.SystemKind = eSyskind.Win32; //Default
        this.Lcid = 1033; //English - United States
        this.LcidInvoke = 1033; //English - United States
        this.CodePage = Encoding.GetEncoding(0).CodePage; //Switched from Default to make it work in Core
        this.MajorVersion = 1361024421;
        this.MinorVersion = 6;
        this.HelpContextID = 0;

        this.Modules.Add(new ExcelVBAModule(this._wb.CodeNameChange)
        {
            Name = "ThisWorkbook",
            Code = "",
            Attributes = GetDocumentAttributes("ThisWorkbook", "0{00020819-0000-0000-C000-000000000046}"),
            Type = eModuleType.Document,
            HelpContext = 0
        });

        foreach (ExcelWorksheet? sheet in this._wb.Worksheets)
        {
            string? name = this.GetModuleNameFromWorksheet(sheet);

            if (!this.Modules.Exists(name))
            {
                this.Modules.Add(new ExcelVBAModule(sheet.CodeNameChange)
                {
                    Name = name,
                    Code = "",
                    Attributes = GetDocumentAttributes(sheet.Name, "0{00020820-0000-0000-C000-000000000046}"),
                    Type = eModuleType.Document,
                    HelpContext = 0
                });
            }
        }

        this._protection = new ExcelVbaProtection(this) { UserProtected = false, HostProtected = false, VbeProtected = false, VisibilityState = true };
    }

    internal string GetModuleNameFromWorksheet(ExcelWorksheet sheet)
    {
        string? name = sheet.Name;
        name = name.Substring(0, name.Length < 31 ? name.Length : 31); //Maximum 31 charachters

        if (this.Modules[name] != null || !ExcelVBAModule.IsValidModuleName(name)) //Check for valid chars, if not valid, set to sheetX.
        {
            int i = sheet.PositionId;
            name = "Sheet" + i.ToString();

            while (this.Modules[name] != null)
            {
                name = "Sheet" + (++i).ToString();
            }
        }

        return name;
    }

    internal static ExcelVbaModuleAttributesCollection GetDocumentAttributes(string name, string clsid)
    {
        ExcelVbaModuleAttributesCollection? attr = new ExcelVbaModuleAttributesCollection();
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Name", Value = name, DataType = eAttributeDataType.String });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Base", Value = clsid, DataType = eAttributeDataType.String });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_GlobalNameSpace", Value = "False", DataType = eAttributeDataType.NonString });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Creatable", Value = "False", DataType = eAttributeDataType.NonString });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_PredeclaredId", Value = "True", DataType = eAttributeDataType.NonString });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Exposed", Value = "False", DataType = eAttributeDataType.NonString });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_TemplateDerived", Value = "False", DataType = eAttributeDataType.NonString });
        attr._list.Add(new ExcelVbaModuleAttribute() { Name = "VB_Customizable", Value = "True", DataType = eAttributeDataType.NonString });

        return attr;
    }

    /// <summary>
    /// Remove the project from the package
    /// </summary>
    public void Remove() => this._wb.RemoveVBAProject();

    internal void RemoveMe()
    {
        if (this.Part == null)
        {
            return;
        }

        foreach (ZipPackageRelationship? rel in this.Part.GetRelationships())
        {
            this._pck.DeleteRelationship(rel.Id);
        }

        if (this._pck.PartExists(this.Uri))
        {
            this._pck.DeletePart(this.Uri);
        }

        this.Part = null;
        this.Modules.Clear();
        this.References.Clear();
        this.Lcid = 0;
        this.LcidInvoke = 0;
        this.CodePage = 0;
        this.MajorVersion = 0;
        this.MinorVersion = 0;
        this.HelpContextID = 0;
    }

    /// <summary>
    /// The name of the project
    /// </summary>
    /// <returns>Returns the name of the project</returns>
    public override string ToString() => this.Name;
}