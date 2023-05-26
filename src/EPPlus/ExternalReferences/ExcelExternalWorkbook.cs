/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Xml;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace OfficeOpenXml.ExternalReferences;

/// <summary>
/// Represents an external workbook.
/// </summary>
public class ExcelExternalWorkbook : ExcelExternalLink
{
    Dictionary<string, int> _sheetNames = new Dictionary<string, int>();
    Dictionary<int, CellStore<object>> _sheetValues = new Dictionary<int, CellStore<object>>();
    Dictionary<int, CellStore<int>> _sheetMetaData = new Dictionary<int, CellStore<int>>();
    Dictionary<int, ExcelExternalNamedItemCollection<ExcelExternalDefinedName>> _definedNamesValues = new Dictionary<int, ExcelExternalNamedItemCollection<ExcelExternalDefinedName>>();
    HashSet<int> _sheetRefresh = new HashSet<int>();
    internal ExcelExternalWorkbook(ExcelWorkbook wb, ExcelPackage p) : base(wb)
    {
        this.CachedWorksheets = new ExcelExternalNamedItemCollection<ExcelExternalWorksheet>();
        this.CachedNames = new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>();
        this.CacheStatus = eExternalWorkbookCacheStatus.NotUpdated;
        this.SetPackage(p, false);
    }
    internal ExcelExternalWorkbook(ExcelWorkbook wb, XmlTextReader reader, ZipPackagePart part, XmlElement workbookElement)  : base(wb, reader, part, workbookElement)
    {
        string? rId = reader.GetAttribute("id", ExcelPackage.schemaRelationships);
        this.Relation = part.GetRelationship(rId);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case "sheetNames":
                        this.ReadSheetNames(reader);
                        break;
                    case "definedNames":
                        this.ReadDefinedNames(reader);
                        break;
                    case "sheetDataSet":
                        this.ReadSheetDataSet(reader, wb);
                        break;
                }
            }
            else if(reader.NodeType==XmlNodeType.EndElement)
            {
                if(reader.Name=="externalBook")
                {
                    reader.Close();
                    break;
                }
            }
        }

        this.CachedWorksheets = new ExcelExternalNamedItemCollection<ExcelExternalWorksheet>();
        this.CachedNames = this.GetNames(-1);
        foreach (string? sheetName in this._sheetNames.Keys)
        {
            int sheetId = this._sheetNames[sheetName];

            this.CachedWorksheets.Add(new ExcelExternalWorksheet(this._sheetValues[sheetId],
                                                                 this._sheetMetaData[sheetId],
                                                                 this._definedNamesValues[sheetId]) 
            { 
                SheetId  = sheetId, 
                Name =sheetName, 
                RefreshError= this._sheetRefresh.Contains(sheetId)
            });
        }

        this.CacheStatus = eExternalWorkbookCacheStatus.LoadedFromPackage;
    }

    /// <summary>
    /// Sets the external link type
    /// </summary>
    public override eExternalLinkType ExternalLinkType
    {
        get
        {
            return eExternalLinkType.ExternalWorkbook;
        }
    }

    private ExcelExternalNamedItemCollection<ExcelExternalDefinedName> GetNames(int ix)
    {
        if(this._definedNamesValues.ContainsKey(ix))
        {
            return this._definedNamesValues[ix];
        }
        else
        {
            return new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>();
        }
    }
    private void ReadSheetDataSet(XmlTextReader reader, ExcelWorkbook wb)
    {
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "sheetDataSet")
            {
                break;
            }
            else if(reader.NodeType == XmlNodeType.Element && reader.Name == "sheetData")
            {
                this.ReadSheetData(reader, wb);
            }
        }
    }
    private void ReadSheetData(XmlTextReader reader, ExcelWorkbook wb)
    {
        int sheetId = int.Parse(reader.GetAttribute("sheetId"));
        if(reader.GetAttribute("refreshError")=="1" && !this._sheetRefresh.Contains(sheetId))
        {
            this._sheetRefresh.Add(sheetId);
        }

        CellStore<object> cellStoreValues = this._sheetValues[sheetId];
        CellStore<int> cellStoreMetaData = this._sheetMetaData[sheetId];

        int row=0, col=0;
        string type="";
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "sheetData")
            {
                break;
            }
            else if(reader.NodeType==XmlNodeType.Element)
            {
                switch(reader.Name)
                {
                    case "row":
                        row = int.Parse(reader.GetAttribute("r"));
                        break;
                    case "cell":
                        ExcelCellBase.GetRowCol(reader.GetAttribute("r"), out row, out col, false);
                        type = reader.GetAttribute("t");
                        string? vm = reader.GetAttribute("vm");
                        if(!string.IsNullOrEmpty(vm))
                        {
                            cellStoreMetaData.SetValue(row, col, int.Parse(vm));
                        }
                        break;
                    case "v":
                        object? v = ConvertUtil.GetValueFromType(reader, type, 0, wb);
                        cellStoreValues.SetValue(row, col, v);
                        break;
                }
            }
        }
    }
    private void ReadDefinedNames(XmlTextReader reader)
    {
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "definedNames")
            {
                break;
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.Name == "definedName")
            {
                int sheetId;
                string? sheetIdAttr = reader.GetAttribute("sheetId");
                if (string.IsNullOrEmpty(sheetIdAttr))
                {
                    sheetId = -1; // -1 represents the workbook level.
                }
                else
                {
                    sheetId = int.Parse(sheetIdAttr);
                }
                    
                ExcelExternalNamedItemCollection<ExcelExternalDefinedName> names = this._definedNamesValues[sheetId];

                string? name = reader.GetAttribute("name");
                names.Add(new ExcelExternalDefinedName() { Name = reader.GetAttribute("name"), RefersTo = reader.GetAttribute("refersTo"), SheetId = sheetId });
            }
        }
    }
    private void ReadSheetNames(XmlTextReader reader)
    {
        int ix = 0;
        this._definedNamesValues.Add(-1, new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>());
        while (reader.Read())
        {
            if(reader.NodeType==XmlNodeType.EndElement && reader.Name== "sheetNames")
            {
                break;
            }
            else if(reader.NodeType==XmlNodeType.Element && reader.Name== "sheetName")
            {
                this._sheetValues.Add(ix, new CellStore<object>());
                this._sheetMetaData.Add(ix, new CellStore<int>());
                this._definedNamesValues.Add(ix, new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>());
                this._sheetNames.Add(reader.GetAttribute("val"), ix++);                    

            }
        }
    }
    /// <summary>
    /// The Uri to the external workbook. This property will be set by the <see cref="File"/> property on save, if it has been set.
    /// </summary>
    public Uri ExternalLinkUri
    {
        get
        {
            return this.Relation?.TargetUri;
        }
        set
        {
            this.Relation.TargetUri = value;
            this._file = null;
        }
    }
    FileInfo _file=null;
    /// <summary>
    /// If the external reference is a file in the filesystem
    /// </summary>
    public FileInfo File
    {
        get
        {
            if(this._file==null)
            {
                string? filePath = this.Relation?.TargetUri?.OriginalString;
                if (string.IsNullOrEmpty(filePath) || HasWebProtocol(filePath))
                {
                    return null;
                }

                if (filePath.StartsWith("file:///"))
                {
                    filePath = filePath.Substring(8);
                }

                try
                {
                        
                    if(this._wb._package.File!=null)
                    {
                        if (string.IsNullOrEmpty(Path.GetDirectoryName(filePath)) || Path.IsPathRooted(filePath) == false)
                        {
                            filePath = this._wb._package.File.DirectoryName + "\\" + filePath;
                        }
                        else
                        {
                            if(Path.IsPathRooted(filePath) == true && filePath[0]==Path.DirectorySeparatorChar)
                            {
                                filePath = this._wb._package.File.Directory.Root.Name + filePath;
                            }
                        }
                    }

                    this._file = new FileInfo(filePath);
                    if(!this._file.Exists && this._wb.ExternalLinks.Directories.Count>0)
                    {
                        this.SetDirectoryIfExists();
                    }
                }
                catch
                {
                    return null;
                }
            }
            return this._file;
        }
        set
        {
            this._file = value;
            if(this._package!=null)
            {
                this._package.File = this.File;
            }
        }
    }

    private void SetDirectoryIfExists()
    {
        foreach(DirectoryInfo? d in this._wb.ExternalLinks.Directories)
        {
            string? file = d.FullName;
            if (file.EndsWith(Path.DirectorySeparatorChar.ToString()) == false)
            {
                file += Path.DirectorySeparatorChar;
            }
            file += this._file.Name;
            if (System.IO.File.Exists(file))
            {
                this._file = new FileInfo(FileHelper.GetRelativeFile(this._wb._package.File, new FileInfo(file)));
                return;
            }
        }
    }

    ExcelPackage _package =null;
    /// <summary>
    /// A reference to the external package, it it has been loaded.
    /// <seealso cref="Load()"/>
    /// </summary>
    public ExcelPackage Package
    {
        get
        {
            return this._package;
        }
    }
    /// <summary>
    /// Tries to Loads the external package using the External Uri into the <see cref="Package"/> property
    /// </summary>
    /// <returns>True if the load succeeded, otherwise false. If false, see <see cref="ExcelExternalLink.ErrorLog"/></returns>
    public bool Load()
    {
        return this.Load(this.File);
    }
    /// <summary>
    /// Tries to Loads the external package using the External Uri into the <see cref="Package"/> property
    /// </summary>
    /// <returns>True if the load succeeded, otherwise false. If false, see <see cref="ExcelExternalLink.ErrorLog"/></returns>
    public bool Load(FileInfo packageFile)
    {
        if (packageFile != null && packageFile.Exists)
        {
            if (!(packageFile.Extension.EndsWith("xlsx", StringComparison.OrdinalIgnoreCase) ||
                  packageFile.Extension.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                  packageFile.Extension.EndsWith(".xlst", StringComparison.OrdinalIgnoreCase)))
            {
                this._errors.Add("EPPlus only supports updating references to files of type xlsx, xlsm and xlst");
                return false;
            }

            this.SetPackage(packageFile);
            return true;
        }

        this._errors.Add($"Loaded file does not exists {packageFile.FullName}");

        return false;
    }
    /// <summary>
    /// Tries to Loads the external package using the External Uri into the <see cref="Package"/> property
    /// </summary>
    /// <returns>True if the load succeeded, otherwise false. If false, see <see cref="ExcelExternalLink.ErrorLog"/> and <see cref="ExcelExternalWorkbook.CacheStatus"/> of each <see cref="ExcelExternalWorkbook"/></returns>
    public bool Load(ExcelPackage package)
    {
        if (package == null || package == this._wb._package)
        {
            this._errors.Add("Load failed. The package can't be null or load itself.");
            return false;
        }

        if (package.File == null)
        {
            this._errors.Add("Load failed. The package must have the File property set to be added as an external reference.");
            return false;
        }

        this.SetPackage(package, true);

        return true;
    }

    private void SetPackage(ExcelPackage package, bool setTarget)
    {
        this._package = package;
        this._package._loadedPackage = this._wb._package;
        this._file = this._package.File;
        if (setTarget)
        {
            this.SetTarget(this._file);
        }
    }
    private void SetPackage(FileInfo file)
    {
        if (this._wb._package.File.Name.Equals(file.Name, StringComparison.CurrentCultureIgnoreCase))
        {
            this._package = this._wb._package;
            return;
        }

        if (this.SetPackageFromOtherReference(this._wb._externalLinks, file) == false)
        {
            this._package = new ExcelPackage(file);
        }

        this._package._loadedPackage = this._wb._package;
        this._file = file;
        this.SetTarget(file);
    }

    private void SetTarget(FileInfo file)
    {
        if (file == null)
        {
            return;
        }

        if (this.IsPathRelative)
        {
            this.Relation.TargetUri = null;
            this.Relation.Target = FileHelper.GetRelativeFile(this._wb._package.File, file, true);
        }
        else
        {
            this.Relation.Target = "file:///" + file.FullName;
            this.Relation.TargetUri = new Uri(this.Relation.Target);
        }
    }

    /// <summary>
    /// If true, sets the path to the workbook as a relative path on <see cref="Load()"/>, if the link is on the same drive.
    /// Otherwise set it as an absolute path. If set to false, the path will always be saved as an absolute path.
    /// If the file path is relative and the file can not be found, the file path will not be updated.
    /// <see cref="Load()"/>
    /// <see cref="File"/>
    /// </summary>
    public bool IsPathRelative { get; set; } = true;
    private bool SetPackageFromOtherReference(ExcelExternalLinksCollection erCollection, FileInfo file)
    {
        if (erCollection == null)
        {
            return false;
        }

        foreach (ExcelExternalLink? er in erCollection)
        {
            if (er!=this && er.ExternalLinkType == eExternalLinkType.ExternalWorkbook)
            {
                ExcelExternalWorkbook? wb = er.As.ExternalWorkbook;
                if (wb._package!=null && wb.File!=null && wb.File.Name.Equals(file.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    this._package=wb._package;
                    return true;
                }

                this.SetPackageFromOtherReference(wb._package?._workbook?._externalLinks, file);
            }
        }
        return false;
    }

    /// <summary>
    /// Updates the external reference cache for the external workbook. To be used a <see cref="Package"/> must be loaded via the <see cref="Load()"/> method.
    /// <seealso cref="CacheStatus"/>
    /// <seealso cref="ExcelExternalLink.ErrorLog"/>
    /// </summary>
    /// <returns>True if the update was successful otherwise false</returns>
    public bool UpdateCache()
    {
        if (this._package == null)
        {
            if (this.Load() == false)
            {
                this.CacheStatus = eExternalWorkbookCacheStatus.Failed;
                this._errors.Add($"Load failed. Can't update cache.");
                return false;
            }
        }

        ILexer? lexer = this._wb.FormulaParser.Lexer;
        this.CachedWorksheets.Clear();
        this.CachedNames.Clear();
        this._definedNamesValues.Clear();
        this._sheetValues.Clear();
        this._sheetMetaData.Clear();
        this._sheetNames.Clear();
        this._definedNamesValues.Add(-1, this.CachedNames);
        foreach (ExcelWorksheet? ws in this._package.Workbook.Worksheets)
        {
            int ix = this.CachedWorksheets.Count;
            this._sheetNames.Add(ws.Name, ix);
            this._sheetValues.Add(ix, new CellStore<object>());
            this._sheetMetaData.Add(ix, new CellStore<int>());
            this._definedNamesValues.Add(ix, new ExcelExternalNamedItemCollection<ExcelExternalDefinedName>());
            this.CachedWorksheets.Add(new ExcelExternalWorksheet(this._sheetValues[ix], this._sheetMetaData[ix], this._definedNamesValues[ix]) { Name = ws.Name, RefreshError = false });
        }

        this.UpdateCacheFromCells();
        this.UpdateCacheFromNames(this._wb, this._wb.Names);
        this.CacheStatus = eExternalWorkbookCacheStatus.Updated;
        return true;
    }

    private void UpdateCacheFromCells()
    {
        foreach (ExcelWorksheet? ws in this._wb.Worksheets)
        {
            CellStoreEnumerator<object>? formulas = new CellStoreEnumerator<object>(ws._formulas);
            foreach (object? f in formulas)
            {
                if (f is int sfIx)
                {
                    ExcelWorksheet.Formulas? sf = ws._sharedFormulas[sfIx];
                    if (sf.Formula.Contains("["))
                    {
                        this.UpdateCacheForFormula(this._wb, sf.Formula, sf.Address);
                    }
                }
                else
                {
                    string? s = f.ToString();
                    if (s.Contains("["))
                    {
                        this.UpdateCacheForFormula(this._wb, s, "");
                    }
                }
            }

            this.UpdateCacheFromNames(this._wb, ws.Names);

            //Update cache for chart references.
            foreach(ExcelDrawing? d in ws.Drawings)
            {
                if(d is ExcelChart c)
                {
                    foreach(ExcelChartSerie? s in c.Series)
                    {
                        if(s.Series.Contains("["))
                        {
                            ExcelAddressBase? a = new ExcelAddressBase(s.Series);
                            if (a.IsExternal)
                            {
                                this.UpdateCacheForAddress(a, "");
                            }
                        }
                        if (s.XSeries.Contains("["))
                        {
                            ExcelAddressBase? a = new ExcelAddressBase(s.XSeries);
                            if (a.IsExternal)
                            {
                                this.UpdateCacheForAddress(a, "");
                            }
                        }
                    }
                }
            }
        }
    }

    private void UpdateCacheFromNames(ExcelWorkbook wb, ExcelNamedRangeCollection names)
    {
        foreach (ExcelNamedRange? n in names)
        {
            if (string.IsNullOrEmpty(n.NameFormula))
            {
                if (n.IsExternal)
                {
                    this.UpdateCacheForAddress(n, "");
                }
            }
            else
            {
                this.UpdateCacheForFormula(wb, n.NameFormula, "");
            }
        }
    }

    /// <summary>
    /// The status of the cache. If the <see cref="UpdateCache" />method fails this status is set to <see cref="eExternalWorkbookCacheStatus.Failed" />
    /// If cache status is set to NotUpdated, the cache will be updated when the package is saved.
    /// <seealso cref="UpdateCache"/>
    /// <seealso cref="ExcelExternalLink.ErrorLog"/>
    /// </summary>
    public eExternalWorkbookCacheStatus CacheStatus { get; private set; }
    private void UpdateCacheForFormula(ExcelWorkbook wb, string formula, string address)
    {
        IEnumerable<Token>? tokens = wb.FormulaParser.Lexer.Tokenize(formula);

        foreach (Token t in tokens)
        {
            if (t.TokenTypeIsSet(TokenType.ExcelAddress) || t.TokenTypeIsSet(TokenType.NameValue))
            {
                if (ExcelCellBase.IsExternalAddress(t.Value))
                {
                    if(t.TokenTypeIsSet(TokenType.ExcelAddress))
                    {
                        ExcelAddressBase a = new ExcelAddressBase(t.Value);
                        int ix = this._wb.ExternalLinks.GetExternalLink(a._wb);
                        if (ix >= 0 && this._wb.ExternalLinks[ix] == this)
                        {
                            this.UpdateCacheForAddress(a, address);
                        }
                    }
                    else
                    {
                        ExcelAddressBase.SplitAddress(t.Value, out string wbRef, out string wsRef, out string nameRef);
                        if (!string.IsNullOrEmpty(wbRef))
                        {
                            int ix = this._wb.ExternalLinks.GetExternalLink(wbRef);
                            if (ix >= 0 && this._wb.ExternalLinks[ix] == this)
                            {
                                string name;
                                if(string.IsNullOrEmpty(wsRef))
                                {
                                    name = nameRef;
                                }
                                else
                                {
                                    name = ExcelCellBase.GetQuotedWorksheetName(wsRef)+"!"+nameRef;
                                }

                                this.UpdateCacheForName(name);
                            }
                        }
                    }
                }
            }
        }
    }

    private void UpdateCacheForName(string name)
    {
        int ix = 0;
        string? wsName = ExcelAddressBase.GetWorksheetPart(name, "", ref ix);
        if (!string.IsNullOrEmpty(wsName))
        {
            name = name.Substring(ix);
        }

        ExcelNamedRange namedRange;
        if (string.IsNullOrEmpty(wsName))
        {
            namedRange = this._package.Workbook.Names.ContainsKey(name) ? this._package.Workbook.Names[name] : null;
        }
        else
        {
            ExcelWorksheet? ws = this._package.Workbook.Worksheets[wsName];
            if (ws == null)
            {
                namedRange = null;
            }
            else
            {
                namedRange = ws.Names.ContainsKey(name) ? ws.Names[name] : null;
            }
        }
        ExcelAddressBase referensTo;
        if(namedRange != null && namedRange._fromRow>0)
        {
            referensTo = new ExcelAddressBase(namedRange.WorkbookLocalAddress);
        }
        else
        {
            referensTo = new ExcelAddressBase("#REF!");
        }

        if(namedRange==null || namedRange.LocalSheetId < 0)
        {
            if (!this.CachedNames.ContainsKey(name))
            {
                this.CachedNames.Add(new ExcelExternalDefinedName() { Name = name, RefersTo = referensTo.Address, SheetId=-1 });
                this.UpdateCacheForAddress(referensTo, "");
            }
        }
        else
        {
            ExcelExternalWorksheet? cws = this.CachedWorksheets[namedRange.LocalSheet.Name];
            if(cws != null)
            {
                if (!cws.CachedNames.ContainsKey(name))
                {
                    cws.CachedNames.Add(new ExcelExternalDefinedName() { Name = name, RefersTo = referensTo.Address, SheetId = namedRange.LocalSheetId });
                    this.UpdateCacheForAddress(referensTo, "");
                }
            }
        }
    }
    private void UpdateCacheForAddress(ExcelAddressBase formulaAddress, string sfAddress)
    {
        if (formulaAddress==null && formulaAddress._fromRow < 0 || formulaAddress._fromCol < 0)
        {
            return;
        }

        if (string.IsNullOrEmpty(sfAddress) == false)
        {
            ExcelAddress? a = new ExcelAddress(sfAddress);
            if (formulaAddress._toColFixed == false)
            {
                formulaAddress._toCol += a.Columns - 1;
                formulaAddress._toRow += a.Rows - 1;
            }
        }

        if (!string.IsNullOrEmpty(formulaAddress.WorkSheetName))
        {
            ExcelWorksheet? ws = this._package.Workbook.Worksheets[formulaAddress.WorkSheetName];
            if (ws == null)
            {
                if (!this.CachedWorksheets.ContainsKey(formulaAddress.WorkSheetName))
                {
                    this.CachedWorksheets.Add(new ExcelExternalWorksheet() { Name = formulaAddress.WorkSheetName, RefreshError = true });
                }
            }
            else
            {
                ExcelExternalWorksheet? cws = this.CachedWorksheets[formulaAddress.WorkSheetName];
                if (cws != null)
                {
                    CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(ws._values, formulaAddress._fromRow, formulaAddress._fromCol, formulaAddress._toRow, formulaAddress._toCol);
                    foreach (ExcelValue v in cse)
                    {
                        cws.CellValues._values.SetValue(cse.Row, cse.Column, v._value);
                    }
                }
            }
        }            
    }

    /// <summary>
    /// String representation
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        if (this.Relation?.TargetUri != null)
        {
            return this.ExternalLinkType.ToString() + "(" + this.Relation.TargetUri.ToString() + ")";
        }
        else
        {
            return base.ToString();
        }
    }
    internal ZipPackageRelationship Relation
    {
        get;
        set;
    }

    /// <summary>
    /// A collection of cached defined names in the external workbook
    /// </summary>
    public ExcelExternalNamedItemCollection<ExcelExternalDefinedName> CachedNames
    {
        get;
    }
    /// <summary>
    /// A collection of cached worksheets in the external workbook
    /// </summary>
    public ExcelExternalNamedItemCollection<ExcelExternalWorksheet> CachedWorksheets
    {
        get;
    }

    internal override void Save(StreamWriter sw)
    {
        if(this.File==null && this.Relation?.TargetUri==null)
        {
            throw new InvalidOperationException($"External reference with Index {this.Index} has no File or Uri set");
        }
        //If sheet names is 0, no update has been performed. Update the cache.
        if(this._sheetNames.Count==0)
        {
            if(this.UpdateCache()==false || this._sheetNames.Count == 0)
            {
                throw (new InvalidDataException($"External reference {this.File.FullName} can't be updated saved. Make sure it contains at least one worksheet. For any errors please check the ErrorLog property of the object after UpdateCache has been called."));
            }
        }

        sw.Write($"<externalBook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{this.Relation.Id}\">");
        sw.Write("<sheetNames>");
        foreach(KeyValuePair<string, int> sheet in this._sheetNames.OrderBy(x=>x.Value))
        {
            sw.Write($"<sheetName val=\"{ConvertUtil.ExcelEscapeString(sheet.Key)}\"/>");
        }
        sw.Write("</sheetNames><definedNames>");
        foreach (int sheet in this._definedNamesValues.Keys)
        {
            foreach (ExcelExternalDefinedName name in this._definedNamesValues[sheet])
            {
                if(name.SheetId<0)
                {
                    sw.Write($"<definedName name=\"{ConvertUtil.ExcelEscapeString(name.Name)}\" refersTo=\"{name.RefersTo}\" />");
                }
                else
                {
                    sw.Write($"<definedName name=\"{ConvertUtil.ExcelEscapeString(name.Name)}\" refersTo=\"{name.RefersTo}\" sheetId=\"{name.SheetId:N0}\"/>");
                }
            }
        }
        sw.Write("</definedNames><sheetDataSet>");
        foreach (int sheetId in this._sheetValues.Keys)
        {
            sw.Write($"<sheetData sheetId=\"{sheetId}\"{(this._sheetRefresh.Contains(sheetId) ? " refreshError=\"1\"" : "")}>");
            CellStoreEnumerator<object>? cellEnum = new CellStoreEnumerator<object>(this._sheetValues[sheetId]);
            CellStore<int>? mdStore = this._sheetMetaData[sheetId];
            int r = -1;
            while(cellEnum.Next())
            {
                if(r!=cellEnum.Row)
                {
                    if(r!=-1)
                    {
                        sw.Write("</row>");
                    }
                    sw.Write($"<row r=\"{cellEnum.Row}\">");                        
                }
                int md=-1;
                if(mdStore.Exists(cellEnum.Row, cellEnum.Column, ref md))
                {
                    sw.Write($"<cell r=\"{ExcelCellBase.GetAddress(cellEnum.Row, cellEnum.Column)}\" md=\"{md}\"{ConvertUtil.GetCellType(cellEnum.Value, true)}><v>{ConvertUtil.ExcelEscapeAndEncodeString(ConvertUtil.GetValueForXml(cellEnum.Value, this._wb.Date1904))}</v></cell>");
                }
                else
                {
                    sw.Write($"<cell r=\"{ExcelCellBase.GetAddress(cellEnum.Row, cellEnum.Column)}\"{ConvertUtil.GetCellType(cellEnum.Value, true)}><v>{ConvertUtil.ExcelEscapeAndEncodeString(ConvertUtil.GetValueForXml(cellEnum.Value, this._wb.Date1904))}</v></cell>");
                }
                r = cellEnum.Row;
            }
            if (r != -1)
            {
                sw.Write("</row>");
            }
            sw.Write("</sheetData>");
        }
        sw.Write("</sheetDataSet></externalBook>");            
    }        
}