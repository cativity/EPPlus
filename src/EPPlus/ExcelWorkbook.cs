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

using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Exporters;
using OfficeOpenXml.Export.HtmlExport.Interfaces;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

#region Public Enum ExcelCalcMode

/// <summary>
/// How the application should calculate formulas in the workbook
/// </summary>
public enum ExcelCalcMode
{
    /// <summary>
    /// Indicates that calculations in the workbook are performed automatically when cell values change. 
    /// The application recalculates those cells that are dependent on other cells that contain changed values. 
    /// This mode of calculation helps to avoid unnecessary calculations.
    /// </summary>
    Automatic,

    /// <summary>
    /// Indicates tables be excluded during automatic calculation
    /// </summary>
    AutomaticNoTable,

    /// <summary>
    /// Indicates that calculations in the workbook be triggered manually by the user. 
    /// </summary>
    Manual
}

#endregion

/// <summary>
/// Represents the Excel workbook and provides access to all the 
/// document properties and worksheets within the workbook.
/// </summary>
public sealed class ExcelWorkbook : XmlHelper, IDisposable
{
    internal class SharedStringItem
    {
        internal int pos;
        internal string Text;
        internal bool isRichText;
    }

    #region Private Properties

    internal ExcelPackage _package;
    internal ExcelWorksheets _worksheets;
    private OfficeProperties _properties;

    private ExcelStyles _styles;

    //internal HashSet<string> _tableSlicerNames = new HashSet<string>();
    internal HashSet<string> _slicerNames;
    internal Dictionary<string, ImageInfo> _images = new Dictionary<string, ImageInfo>();

    internal bool GetPivotCacheFromAddress(string fullAddress, out PivotTableCacheInternal cacheReference)
    {
        if (this._pivotTableCaches.TryGetValue(fullAddress, out PivotTableCacheRangeInfo cacheInfo))
        {
            cacheReference = cacheInfo.PivotCaches[0];

            return true;
        }

        cacheReference = null;

        return false;
    }

    internal void LoadAllDrawings(string loadingWsName)
    {
        if (this._worksheets._areDrawingsLoaded)
        {
            return;
        }

        this._worksheets._areDrawingsLoaded = true;

        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (loadingWsName.Equals(ws.Name, StringComparison.OrdinalIgnoreCase) == false)
            {
                ws.LoadDrawings();
            }
        }
    }

    internal string GetSlicerName(string name)
    {
        if (this._slicerNames == null)
        {
            this.LoadSlicerNames();
        }

        return GetUniqueName(name, this._slicerNames);
    }

    internal bool CheckSlicerNameIsUnique(string name)
    {
        if (this._slicerNames == null)
        {
            this.LoadSlicerNames();
        }

        if (this._slicerNames.Contains(name))
        {
            return false;
        }

        _ = this._slicerNames.Add(name);

        return true;
    }

    private void LoadSlicerNames()
    {
        this._slicerNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            foreach (ExcelDrawing d in ws.Drawings)
            {
                if (d is ExcelTableSlicer || d is ExcelPivotTableSlicer)
                {
                    _ = this._slicerNames.Add(d.Name);
                }
            }
        }
    }

    private static string GetUniqueName(string name, HashSet<string> hs)
    {
        string n = name;
        int ix = 1;

        while (hs.Contains(n))
        {
            n = name + $"{ix++}";
        }

        return n;
    }

    #endregion

    #region ExcelWorkbook Constructor

    /// <summary>
    /// Creates a new instance of the ExcelWorkbook class.
    /// </summary>
    /// <param name="package">The parent package</param>
    /// <param name="namespaceManager">NamespaceManager</param>
    internal ExcelWorkbook(ExcelPackage package, XmlNamespaceManager namespaceManager)
        : base(namespaceManager)
    {
        this._package = package;
        this.SetUris();

        this._names = new ExcelNamedRangeCollection(this);
        this._namespaceManager = namespaceManager;
        this.TopNode = this.WorkbookXml.DocumentElement;

        this.SchemaNodeOrder = new string[]
        {
            "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "functionPrototypes",
            "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
            "fileRecoveryPr", "webPublishObjects", "extLst"
        };

        this.FullCalcOnLoad = true; //Full calculation on load by default, for both new workbooks and templates.
        this.GetSharedStrings();
    }

    /// <summary>
    /// Load all pivot cache ids and there uri's
    /// </summary>
    internal void LoadPivotTableCaches()
    {
        XmlNodeList? pts = this.GetNodes("d:pivotCaches/d:pivotCache");

        if (pts != null)
        {
            foreach (XmlElement pt in pts)
            {
                string rid = pt.GetAttribute("r:id");
                string cacheId = pt.GetAttribute("cacheId");
                ZipPackageRelationship rel = this.Part.GetRelationship(rid);
                this._pivotTableIds.Add(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri), int.Parse(cacheId));
            }
        }
    }

    private void SetUris()
    {
        foreach (ZipPackageRelationship rel in this._package.ZipPackage.GetRelationships())
        {
            if (rel.RelationshipType == ExcelPackage.schemaRelationships + "/officeDocument")
            {
                this.WorkbookUri = rel.TargetUri;

                break;
            }
        }

        if (this.WorkbookUri == null)
        {
            this.WorkbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
        }
        else
        {
            foreach (ZipPackageRelationship rel in this.Part.GetRelationships())
            {
                switch (rel.RelationshipType)
                {
                    case ExcelPackage.schemaRelationships + "/sharedStrings":
                        this.SharedStringsUri = UriHelper.ResolvePartUri(this.WorkbookUri, rel.TargetUri);

                        break;

                    case ExcelPackage.schemaRelationships + "/styles":
                        this.StylesUri = UriHelper.ResolvePartUri(this.WorkbookUri, rel.TargetUri);

                        break;

                    case ExcelPackage.schemaPersonsRelationShips:
                        this.PersonsUri = UriHelper.ResolvePartUri(this.WorkbookUri, rel.TargetUri);

                        break;
                }
            }
        }

        if (this.SharedStringsUri == null)
        {
            this.SharedStringsUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
        }

        if (this.StylesUri == null)
        {
            this.StylesUri = new Uri("/xl/styles.xml", UriKind.Relative);
        }

        if (this.PersonsUri == null)
        {
            this.PersonsUri = new Uri("/xl/persons/person.xml", UriKind.Relative);
        }
    }

    #endregion

    internal Dictionary<string, SharedStringItem> _sharedStrings = new Dictionary<string, SharedStringItem>(); //Used when reading cells.
    internal List<SharedStringItem> _sharedStringsList = new List<SharedStringItem>(); //Used when reading cells.
    internal ExcelNamedRangeCollection _names;
    internal int _nextDrawingId = 2;
    internal int _nextTableID = int.MinValue;
    internal int _nextPivotCacheId = 1;

    internal int GetNewPivotCacheId() => this._nextPivotCacheId++;

    internal void SetNewPivotCacheId(int value)
    {
        if (value >= this._nextPivotCacheId)
        {
            this._nextPivotCacheId = value + 1;
        }
    }

    internal int _nextPivotTableID = int.MinValue;
    internal XmlNamespaceManager _namespaceManager;

    internal FormulaParser _formulaParser;
    internal ExcelThreadedCommentPersonCollection _threadedCommentPersons;
    internal FormulaParserManager _parserManager;
    internal CellStore<List<Token>> _formulaTokens;

    internal class PivotTableCacheRangeInfo
    {
        public string Address { get; set; }

        public List<PivotTableCacheInternal> PivotCaches { get; set; }
    }

    internal Dictionary<string, PivotTableCacheRangeInfo> _pivotTableCaches = new Dictionary<string, PivotTableCacheRangeInfo>();
    internal Dictionary<Uri, int> _pivotTableIds = new Dictionary<Uri, int>();

    /// <summary>
    /// Read shared strings to list
    /// </summary>
    private void GetSharedStrings()
    {
        if (this._package.ZipPackage.PartExists(this.SharedStringsUri))
        {
            XmlDocument xml = this._package.GetXmlFromUri(this.SharedStringsUri);
            XmlNodeList nl = xml.SelectNodes("//d:sst/d:si", this.NameSpaceManager);
            this._sharedStringsList = new List<SharedStringItem>();

            if (nl != null)
            {
                foreach (XmlNode node in nl)
                {
                    XmlNode n = node.SelectSingleNode("d:t", this.NameSpaceManager);

                    if (n != null)
                    {
                        this._sharedStringsList.Add(new SharedStringItem() { Text = ConvertUtil.ExcelDecodeString(n.InnerText) });
                    }
                    else
                    {
                        this._sharedStringsList.Add(new SharedStringItem() { Text = node.InnerXml, isRichText = true });
                    }
                }
            }

            //Delete the shared string part, it will be recreated when the package is saved.
            foreach (ZipPackageRelationship rel in this.Part.GetRelationships())
            {
                if (rel.TargetUri.OriginalString.EndsWith("sharedstrings.xml", StringComparison.OrdinalIgnoreCase))
                {
                    this.Part.DeleteRelationship(rel.Id);

                    break;
                }
            }

            this._package.ZipPackage.DeletePart(this.SharedStringsUri); //Remove the part, it is recreated when saved.
        }
    }

    internal void GetDefinedNames()
    {
        XmlNodeList nl = this.WorkbookXml.SelectNodes("//d:definedNames/d:definedName", this.NameSpaceManager);

        if (nl != null)
        {
            foreach (XmlElement elem in nl)
            {
                string fullAddress = elem.InnerText.TrimStart().TrimEnd();

                ExcelWorksheet nameWorksheet;

                if (!int.TryParse(elem.GetAttribute("localSheetId"), NumberStyles.Number, CultureInfo.InvariantCulture, out int localSheetID))
                {
                    localSheetID = -1;
                    nameWorksheet = null;
                }
                else
                {
                    nameWorksheet = this.Worksheets[localSheetID + this._package._worksheetAdd];
                }

                ExcelAddressBase.AddressType addressType = ExcelAddressBase.IsValid(fullAddress);
                ExcelNamedRange namedRange;

                if (addressType == ExcelAddressBase.AddressType.Invalid
                    || addressType == ExcelAddressBase.AddressType.InternalName
                    || addressType == ExcelAddressBase.AddressType.ExternalName
                    || addressType == ExcelAddressBase.AddressType.Formula
                    || addressType == ExcelAddressBase.AddressType.ExternalAddress) //A value or a formula
                {
                    namedRange = this.AddFormulaOrValueName(elem, fullAddress, nameWorksheet);
                }
                else
                {
                    ExcelAddress addr = new ExcelAddress(fullAddress, this._package, null);

                    if (addr._fromRow <= 0
                        && fullAddress.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) < 0) // Address is not valid, add as a formula instead
                    {
                        namedRange = this.AddFormulaOrValueName(elem, fullAddress, nameWorksheet);
                    }
                    else if (localSheetID > -1)
                    {
                        if (string.IsNullOrEmpty(addr._ws))
                        {
                            ExcelRangeBase addressRange =
                                this.CreateRangeForName(this.Worksheets[localSheetID + this._package._worksheetAdd],
                                                        fullAddress,
                                                        out bool allowRelativeAddress);

                            namedRange = this.Worksheets[localSheetID + this._package._worksheetAdd]
                                             .Names.AddName(elem.GetAttribute("name"), addressRange, allowRelativeAddress);
                        }
                        else
                        {
                            ExcelRangeBase addressRange = this.CreateRangeForName(this.Worksheets[addr._ws], fullAddress, out bool allowRelativeAddress);

                            namedRange = this.Worksheets[localSheetID + this._package._worksheetAdd]
                                             .Names.AddName(elem.GetAttribute("name"), addressRange, allowRelativeAddress);
                        }
                    }
                    else
                    {
                        ExcelWorksheet? ws = this.Worksheets[addr._ws];

                        if (ws == null)
                        {
                            namedRange = this._names.AddFormula(elem.GetAttribute("name"), fullAddress);
                        }
                        else
                        {
                            ExcelRangeBase addressRange = this.CreateRangeForName(ws, fullAddress, out bool allowRelativeAddress);
                            namedRange = this._names.AddName(elem.GetAttribute("name"), addressRange, allowRelativeAddress);
                        }
                    }
                }

                if (elem.GetAttribute("hidden") == "1" && namedRange != null)
                {
                    namedRange.IsNameHidden = true;
                }

                if (!string.IsNullOrEmpty(elem.GetAttribute("comment")))
                {
                    namedRange.NameComment = elem.GetAttribute("comment");
                }
            }
        }
    }

    private ExcelNamedRange AddFormulaOrValueName(XmlElement elem, string fullAddress, ExcelWorksheet nameWorksheet)
    {
        ExcelNamedRange namedRange;
        ExcelRangeBase range = new ExcelRangeBase(this, nameWorksheet, elem.GetAttribute("name"), true);

        if (nameWorksheet == null)
        {
            namedRange = this._names.AddName(elem.GetAttribute("name"), range);
        }
        else
        {
            namedRange = nameWorksheet.Names.AddName(elem.GetAttribute("name"), range);
        }

        if (ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "\"")) //String value
        {
            namedRange.NameValue = fullAddress.Substring(1, fullAddress.Length - 2);
        }
        else if (double.TryParse(fullAddress, NumberStyles.Number, CultureInfo.InvariantCulture, out double value))
        {
            namedRange.NameValue = value;
        }
        else
        {
            namedRange.NameFormula = fullAddress;
        }

        return namedRange;
    }

    private ExcelRangeBase CreateRangeForName(ExcelWorksheet worksheet, string fullAddress, out bool allowRelativeAddress)
    {
        bool iR = false;
        ExcelRangeBase range = new ExcelRangeBase(this, worksheet, fullAddress, false);
        ExcelAddressBase addr = range.ToInternalAddress();

        if (addr._fromColFixed || addr._toColFixed || addr._fromRowFixed || addr._toRowFixed)
        {
            iR = true;
        }

        allowRelativeAddress = iR;

        return range;
    }

    internal void RemoveSlicerCacheReference(string relId, eSlicerSourceType sourceType)
    {
        string path;

        if (sourceType == eSlicerSourceType.PivotTable)
        {
            path = $"d:extLst/d:ext/x14:slicerCaches/x14:slicerCache[@r:id='{relId}']";
        }
        else
        {
            path = $"d:extLst/d:ext/x15:slicerCaches/x14:slicerCache[@r:id='{relId}']";
        }

        XmlNode? node = this.GetNode(path);

        if (node != null)
        {
            if (node.ParentNode.ChildNodes.Count > 1)
            {
                _ = node.ParentNode.RemoveChild(node);
            }
            else
            {
                _ = node.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode);
            }
        }
    }

    internal ExcelRangeBase GetRange(ExcelWorksheet ws, string function)
    {
        switch (ExcelAddressBase.IsValid(function))
        {
            case ExcelAddressBase.AddressType.InternalAddress:
                ExcelAddress addr = new ExcelAddress(function);

                if (string.IsNullOrEmpty(addr.WorkSheetName))
                {
                    return ws.Cells[function];
                }
                else
                {
                    ExcelWorksheet? otherWs = this.Worksheets[addr.WorkSheetName];

                    if (otherWs == null)
                    {
                        return null;
                    }
                    else
                    {
                        return otherWs.Cells[addr.Address];
                    }
                }

            case ExcelAddressBase.AddressType.InternalName:
                if (this.Names.ContainsKey(function))
                {
                    return this.Names[function];
                }
                else if (ws.Names.ContainsKey(function))
                {
                    return ws.Names[function];
                }
                else if (ws.Tables[function] != null)
                {
                    return ws.Cells[ws.Tables[function].Address.Address];
                }
                else
                {
                    ExcelAddress nameAddr = new ExcelAddress(function);

                    if (string.IsNullOrEmpty(nameAddr.WorkSheetName))
                    {
                        return null;
                    }
                    else
                    {
                        ExcelWorksheet? otherWs = this.Worksheets[nameAddr.WorkSheetName];

                        if (otherWs != null && otherWs.Names.ContainsKey(nameAddr.Address))
                        {
                            return otherWs.Names[nameAddr.Address];
                        }

                        return null;
                    }
                }

            case ExcelAddressBase.AddressType.Formula:
                return null;

            default:
                return null;
        }
    }

    internal int GetPivotCacheId(Uri cacheDefinitionUri)
    {
        foreach (ZipPackageRelationship rel in this.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheDefinition"))
        {
            if (cacheDefinitionUri == UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri))
            {
                return this.GetXmlNodeInt($"d:pivotCaches/d:pivotCache[@r:id='{rel.Id}']/@cacheId");
            }
        }

        return int.MinValue;
    }

    #region Worksheets

    /// <summary>
    /// Provides access to all the worksheets in the workbook.
    /// Note: Worksheets index either starts by 0 or 1 depending on the Excelpackage.Compatibility.IsWorksheets1Based property.
    /// Default is 1 for .Net 3.5 and .Net 4 and 0 for .Net Core.
    /// </summary>
    public ExcelWorksheets Worksheets
    {
        get
        {
            if (this._worksheets == null)
            {
                XmlNode? sheetsNode = this._workbookXml.DocumentElement.SelectSingleNode("d:sheets", this._namespaceManager) ?? this.CreateNode("d:sheets");

                this._worksheets = new ExcelWorksheets(this._package, this._namespaceManager, sheetsNode);
            }

            return this._worksheets;
        }
    }

    #endregion

    /// <summary>
    /// Create an html exporter for the supplied ranges.
    /// </summary>
    /// <param name="ranges">The ranges to create the report from. All ranges must originate from the current workbook. </param>
    /// <returns>The HTML exporter.</returns>
    /// <exception cref="InvalidOperationException"></exception>
    public IExcelHtmlRangeExporter CreateHtmlExporter(params ExcelRangeBase[] ranges)
    {
        foreach (ExcelRangeBase range in ranges)
        {
            if (range._workbook != this)
            {
                throw new InvalidOperationException("All ranges must come from the current workbook");
            }
        }

        return new ExcelHtmlWorkbookExporter(ranges);
    }

    //public ExcelHtmlRangeExporter CreateHtmlExporter(params ExcelRangeBase[] ranges)
    //{
    //    foreach (var range in ranges)
    //    {
    //        if (range._workbook != this)
    //        {
    //            throw new InvalidOperationException("All ranges must come from the current workbook");
    //        }
    //    }
    //    return new Export.HtmlExport.ExcelHtmlRangeExporter(ranges);
    //}
    /// <summary>
    /// Provides access to named ranges
    /// </summary>
    public ExcelNamedRangeCollection Names => this._names;

    internal ExcelExternalLinksCollection _externalLinks;

    /// <summary>
    /// A collection of links to external workbooks and it's cached data.
    /// This collection can also contain DDE and OLE links. DDE and OLE are readonly and cannot be added.
    /// </summary>
    public ExcelExternalLinksCollection ExternalLinks => this._externalLinks ??= new ExcelExternalLinksCollection(this);

    #region Workbook Properties

    decimal _standardFontWidth = decimal.MinValue;
    string _fontID = "";

    internal FormulaParser FormulaParser => this._formulaParser ??= new FormulaParser(new EpplusExcelDataProvider(this._package));

    /// <summary>
    /// Manage the formula parser.
    /// Add your own functions or replace native ones, parse formulas or attach a logger.
    /// </summary>
    public FormulaParserManager FormulaParserManager => this._parserManager ??= new FormulaParserManager(this.FormulaParser);

    /// <summary>
    /// Represents a collection of <see cref="ExcelThreadedCommentPerson"/>s in the workbook.
    /// </summary>
    public ExcelThreadedCommentPersonCollection ThreadedCommentPersons => this._threadedCommentPersons ??= new ExcelThreadedCommentPersonCollection(this);

    /// <summary>
    /// Max font width for the workbook
    /// <remarks>This method uses GDI. If you use Azure or another environment that does not support GDI, you have to set this value manually if you don't use the standard Calibri font</remarks>
    /// </summary>
    public decimal MaxFontWidth
    {
        get
        {
            int ix = this.Styles.GetNormalStyleIndex();

            if (ix >= 0)
            {
                ExcelFont font = this.Styles.NamedStyles[ix].Style.Font;

                if (font.Index == int.MinValue)
                {
                    font.Index = 0;
                }

                if (this._standardFontWidth == decimal.MinValue || this._fontID != font.Id)
                {
                    try
                    {
                        this._standardFontWidth = FontSize.GetWidthPixels(font.Name, font.Size);
                        this._fontID = this.Styles.NamedStyles[ix].Style.Font.Id;
                    }
                    catch //Error, Font missing and Calibri removed in dictionary
                    {
                        this._standardFontWidth = (int)(font.Size * (2D / 3D)); //Aprox for Calibri.
                    }
                }
            }
            else
            {
                this._standardFontWidth = 7; //Calibri 11
            }

            return this._standardFontWidth;
        }
        set => this._standardFontWidth = value;
    }

    internal static decimal GetHeightPixels(string fontName, float fontSize)
    {
        Dictionary<float, short> font = FontSize.GetFontSize(fontName, false);

        if (font.ContainsKey(fontSize))
        {
            return Convert.ToDecimal(font[fontSize]);
        }
        else
        {
            float min = -1;

            foreach (float size in font.Keys)
            {
                if (min < size && size > fontSize)
                {
                    if (min == -1)
                    {
                        min = size;
                    }

                    break;
                }

                min = size;
            }

            if (min > -1)
            {
                return font[min];
            }

            return 20; //Default pixels, Calibri 11
        }
    }

    ExcelProtection _protection;

    /// <summary>
    /// Access properties to protect or unprotect a workbook
    /// </summary>
    public ExcelProtection Protection
    {
        get
        {
            if (this._protection == null)
            {
                this._protection = new ExcelProtection(this.NameSpaceManager, this.TopNode, this);
                this._protection.SchemaNodeOrder = this.SchemaNodeOrder;
            }

            return this._protection;
        }
    }

    ExcelWorkbookView _view;

    /// <summary>
    /// Access to workbook view properties
    /// </summary>
    public ExcelWorkbookView View => this._view ??= new ExcelWorkbookView(this.NameSpaceManager, this.TopNode, this);

    ExcelVbaProject _vba;

    /// <summary>
    /// A reference to the VBA project.
    /// Null if no project exists.
    /// Use Workbook.CreateVBAProject to create a new VBA-Project
    /// </summary>
    public ExcelVbaProject VbaProject
    {
        get
        {
            if (this._vba == null)
            {
                if (this._package.ZipPackage.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
                {
                    this._vba = new ExcelVbaProject(this);
                }
                else if (this.Part.ContentType == ContentTypes.contentTypeWorkbookMacroEnabled) //Project is macro enabled, but no bin file exists.
                {
                    this.CreateVBAProject();
                }
            }

            return this._vba;
        }
    }

    /// <summary>
    /// Remove the from the file VBA project.
    /// </summary>
    public void RemoveVBAProject()
    {
        if (this._vba != null)
        {
            this._vba.RemoveMe();
            this.Part.ContentType = ContentTypes.contentTypeWorkbookDefault;
            this._vba = null;
        }
    }

    /// <summary>
    /// Create an empty VBA project.
    /// </summary>
    public void CreateVBAProject()
    {
        if (this._vba != null || this._package.ZipPackage.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
        {
            throw new InvalidOperationException("VBA project already exists.");
        }

        this.Part.ContentType = ContentTypes.contentTypeWorkbookMacroEnabled;
        this._vba = new ExcelVbaProject(this);
        this._vba.Create();
    }

    /// <summary>
    /// URI to the workbook inside the package
    /// </summary>
    internal Uri WorkbookUri { get; private set; }

    /// <summary>
    /// URI to the styles inside the package
    /// </summary>
    internal Uri StylesUri { get; private set; }

    /// <summary>
    /// URI to the shared strings inside the package
    /// </summary>
    internal Uri SharedStringsUri { get; private set; }

    /// <summary>
    /// URI to the person elements inside the package
    /// </summary>
    internal Uri PersonsUri { get; private set; }

    /// <summary>
    /// Returns a reference to the workbook's part within the package
    /// </summary>
    internal ZipPackagePart Part => this._package.ZipPackage.GetPart(this.WorkbookUri);

    #region WorkbookXml

    private XmlDocument _workbookXml;

    /// <summary>
    /// Provides access to the XML data representing the workbook in the package.
    /// </summary>
    public XmlDocument WorkbookXml
    {
        get
        {
            if (this._workbookXml == null)
            {
                this.CreateWorkbookXml(this._namespaceManager);
            }

            return this._workbookXml;
        }
    }

    const string codeModuleNamePath = "d:workbookPr/@codeName";

    internal string CodeModuleName
    {
        get => this.GetXmlNodeString(codeModuleNamePath);
        set => this.SetXmlNodeString(codeModuleNamePath, value);
    }

    internal void CodeNameChange(string value) => this.CodeModuleName = value;

    /// <summary>
    /// The VBA code module if the package has a VBA project. Otherwise this propery is null.
    /// <seealso cref="CreateVBAProject"/>
    /// </summary>
    public ExcelVBAModule CodeModule
    {
        get
        {
            if (this.VbaProject != null)
            {
                return this.VbaProject.Modules[this.CodeModuleName];
            }
            else
            {
                return null;
            }
        }
    }

    const string date1904Path = "d:workbookPr/@date1904";
    internal const double date1904Offset = 365.5 * 4; // offset to fix 1900 and 1904 differences, 4 OLE years
    private bool? date1904Cache;

    internal bool ExistsPivotCache(int cacheID, ref int newID)
    {
        newID = cacheID;
        bool ret = true;

        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (ws is ExcelChartsheet)
            {
                continue;
            }

            foreach (ExcelPivotTable pt in ws.PivotTables)
            {
                if (pt.CacheId == cacheID)
                {
                    ret = false;
                }

                if (pt.CacheId >= newID)
                {
                    newID = pt.CacheId + 1;
                }
            }
        }

        if (ret)
        {
            newID = cacheID; //Not Found, return same ID
        }

        return ret;
    }

    /// <summary>
    /// The date systems used by Microsoft Excel can be based on one of two different dates. By default, a serial number of 1 in Microsoft Excel represents January 1, 1900.
    /// The default for the serial number 1 can be changed to represent January 2, 1904.
    /// This option was included in Microsoft Excel for Windows to make it compatible with Excel for the Macintosh, which defaults to January 2, 1904.
    /// </summary>
    public bool Date1904
    {
        get
        {
            this.date1904Cache ??= this.GetXmlNodeBool(date1904Path, false);

            return this.date1904Cache.Value;
        }
        set
        {
            if (this.Date1904 != value)
            {
                // Like Excel when the option it's changed update it all cells with Date format
                foreach (ExcelWorksheet ws in this.Worksheets)
                {
                    if (ws is ExcelChartsheet)
                    {
                        continue;
                    }

                    ws.UpdateCellsWithDate1904Setting();
                }
            }

            this.date1904Cache = value;
            this.SetXmlNodeBool(date1904Path, value, false);
        }
    }

    /// <summary>
    /// Create or read the XML for the workbook.
    /// </summary>
    private void CreateWorkbookXml(XmlNamespaceManager namespaceManager)
    {
        if (this._package.ZipPackage.PartExists(this.WorkbookUri))
        {
            this._workbookXml = this._package.GetXmlFromUri(this.WorkbookUri);
        }
        else
        {
            // create a new workbook part and add to the package
            ZipPackagePart partWorkbook = this._package.ZipPackage.CreatePart(this.WorkbookUri,
                                                                              @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                                                                              this._package.Compression);

            // create the workbook
            this._workbookXml = new XmlDocument(namespaceManager.NameTable);

            this._workbookXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;

            // create the workbook element
            XmlElement wbElem = this._workbookXml.CreateElement("workbook", ExcelPackage.schemaMain);

            // Add the relationships namespace
            wbElem.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);

            _ = this._workbookXml.AppendChild(wbElem);

            // create the bookViews and workbooks element
            XmlElement bookViews = this._workbookXml.CreateElement("bookViews", ExcelPackage.schemaMain);
            _ = wbElem.AppendChild(bookViews);
            XmlElement workbookView = this._workbookXml.CreateElement("workbookView", ExcelPackage.schemaMain);
            _ = bookViews.AppendChild(workbookView);

            // save it to the package
            StreamWriter stream = new StreamWriter(partWorkbook.GetStream(FileMode.Create, FileAccess.Write));
            this._workbookXml.Save(stream);

            //stream.Close();
            ZipPackage.Flush();
        }
    }

    #endregion

    #region StylesXml

    private XmlDocument _stylesXml;

    /// <summary>
    /// Provides access to the XML data representing the styles in the package. 
    /// </summary>
    public XmlDocument StylesXml
    {
        get
        {
            if (this._stylesXml == null)
            {
                if (this._package.ZipPackage.PartExists(this.StylesUri))
                {
                    this._stylesXml = this._package.GetXmlFromUri(this.StylesUri);
                }
                else
                {
                    // create a new styles part and add to the package
                    ZipPackagePart part = this._package.ZipPackage.CreatePart(this.StylesUri,
                                                                              @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                                                                              this._package.Compression);

                    // create the style sheet

                    StringBuilder xml = new StringBuilder("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                    _ = xml.Append("<numFmts />");
                    _ = xml.Append("<fonts count=\"1\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts>");
                    _ = xml.Append("<fills><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills>");
                    _ = xml.Append("<borders><border><left /><right /><top /><bottom /><diagonal /></border></borders>");
                    _ = xml.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs>");
                    _ = xml.Append("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" xfId=\"0\" /></cellXfs>");
                    _ = xml.Append("<cellStyles><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles>");
                    _ = xml.Append("<dxfs count=\"0\" />");
                    _ = xml.Append("</styleSheet>");

                    this._stylesXml = new XmlDocument();
                    this._stylesXml.LoadXml(xml.ToString());

                    //Save it to the package
                    StreamWriter stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));

                    this._stylesXml.Save(stream);

                    //stream.Close();
                    ZipPackage.Flush();

                    // create the relationship between the workbook and the new shared strings part
                    _ = this._package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(this.WorkbookUri, this.StylesUri),
                                                                   TargetMode.Internal,
                                                                   ExcelPackage.schemaRelationships + "/styles");

                    ZipPackage.Flush();
                }
            }

            return this._stylesXml;
        }
        set => this._stylesXml = value;
    }

    /// <summary>
    /// Package styles collection. Used internally to access style data.
    /// </summary>
    public ExcelStyles Styles => this._styles ??= new ExcelStyles(this.NameSpaceManager, this.StylesXml, this);

    #endregion

    #region Office Document Properties

    /// <summary>
    /// The office document properties
    /// </summary>
    public OfficeProperties Properties =>

        //  Create a NamespaceManager to handle the default namespace, 
        //  and create a prefix for the default namespace:                   
        this._properties ??= new OfficeProperties(this._package, this.NameSpaceManager);

    #endregion

    #region CalcMode

    private string CALC_MODE_PATH = "d:calcPr/@calcMode";

    /// <summary>
    /// Calculation mode for the workbook.
    /// </summary>
    public ExcelCalcMode CalcMode
    {
        get
        {
            string calcMode = this.GetXmlNodeString(this.CALC_MODE_PATH);

            switch (calcMode)
            {
                case "autoNoTable":
                    return ExcelCalcMode.AutomaticNoTable;

                case "manual":
                    return ExcelCalcMode.Manual;

                default:
                    return ExcelCalcMode.Automatic;
            }
        }
        set
        {
            switch (value)
            {
                case ExcelCalcMode.AutomaticNoTable:
                    this.SetXmlNodeString(this.CALC_MODE_PATH, "autoNoTable");

                    break;

                case ExcelCalcMode.Manual:
                    this.SetXmlNodeString(this.CALC_MODE_PATH, "manual");

                    break;

                default:
                    this.SetXmlNodeString(this.CALC_MODE_PATH, "auto");

                    break;
            }
        }

        #endregion
    }

    private const string FULL_CALC_ON_LOAD_PATH = "d:calcPr/@fullCalcOnLoad";

    /// <summary>
    /// Should Excel do a full calculation after the workbook has been loaded?
    /// <remarks>This property is always true for both new workbooks and loaded templates(on load). If this is not the wanted behavior set this property to false.</remarks>
    /// </summary>
    public bool FullCalcOnLoad
    {
        get => this.GetXmlNodeBool(FULL_CALC_ON_LOAD_PATH);
        set => this.SetXmlNodeBool(FULL_CALC_ON_LOAD_PATH, value);
    }

    ExcelThemeManager _theme;

    /// <summary>
    /// Create and manage the theme for the workbook.
    /// </summary>
    public ExcelThemeManager ThemeManager => this._theme ??= new ExcelThemeManager(this);

    const string defaultThemeVersionPath = "d:workbookPr/@defaultThemeVersion";

    /// <summary>
    /// The default version of themes to apply in the workbook
    /// </summary>
    public int? DefaultThemeVersion
    {
        get => this.GetXmlNodeIntNull(defaultThemeVersionPath);
        set
        {
            if (value is null)
            {
                this.DeleteNode(defaultThemeVersionPath);
            }
            else
            {
                this.SetXmlNodeString(defaultThemeVersionPath, value.ToString());
            }
        }
    }

    #endregion

    #region Workbook Private Methods

    #region Save // Workbook Save

    /// <summary>
    /// Saves the workbook and all its components to the package.
    /// For internal use only!
    /// </summary>
    internal void Save() // Workbook Save
    {
        if (this.Worksheets.Count == 0)
        {
            throw new InvalidOperationException("The workbook must contain at least one worksheet");
        }

        this.DeleteCalcChain();

        if (this._vba == null && !this._package.ZipPackage.PartExists(new Uri(ExcelVbaProject.PartUri, UriKind.Relative)))
        {
            if (this.Part.ContentType != ContentTypes.contentTypeWorkbookDefault && this.Part.ContentType != ContentTypes.contentTypeWorkbookMacroEnabled)
            {
                this.Part.ContentType = ContentTypes.contentTypeWorkbookDefault;
            }
        }
        else
        {
            if (this.Part.ContentType != ContentTypes.contentTypeWorkbookMacroEnabled)
            {
                this.Part.ContentType = ContentTypes.contentTypeWorkbookMacroEnabled;
            }
        }

        this.UpdateDefinedNamesXml();

        if (this.HasLoadedPivotTables)
        {
            //Updates the Workbook Xml, so must be before saving the wookbook part 
            this.SavePivotTableCaches();
        }

        if (this._externalLinks != null)
        {
            this.SaveExternalLinks();
        }

        // save the workbook
        if (this._workbookXml != null)
        {
            if (this.Worksheets[this._package._worksheetAdd].Hidden != eWorkSheetHidden.Visible)
            {
                int? ix = this.Worksheets.GetFirstVisibleSheetIndex();

                if (ix > this.View.FirstSheet)
                {
                    this.View.FirstSheet = ix;
                }
            }

            this._package.SavePart(this.WorkbookUri, this._workbookXml);
        }

        // save the properties of the workbook
        if (this._properties != null)
        {
            this._properties.Save();
        }

        //Save the Theme
        this.ThemeManager.Save();

        // save the style sheet
        this.Styles.UpdateXml();
        this._package.SavePart(this.StylesUri, this.StylesXml);

        // save persons
        this._threadedCommentPersons?.Save(this._package, this.Part, this.PersonsUri);

        // save threaded comments

        // save all the open worksheets
        bool isProtected = this.Protection.LockWindows || this.Protection.LockStructure;

        foreach (ExcelWorksheet worksheet in this.Worksheets)
        {
            if (isProtected && this.Protection.LockWindows)
            {
                worksheet.View.WindowProtection = true;
            }

            worksheet.Save();
            worksheet.Part.SaveHandler = worksheet.SaveHandler;
        }

        // Issue 15252: save SharedStrings only once
        ZipPackagePart part;

        if (this._package.ZipPackage.PartExists(this.SharedStringsUri))
        {
            part = this._package.ZipPackage.GetPart(this.SharedStringsUri);
        }
        else
        {
            part = this._package.ZipPackage.CreatePart(this.SharedStringsUri,
                                                       @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                                                       this._package.Compression);

            _ = this.Part.CreateRelationship(UriHelper.GetRelativeUri(this.WorkbookUri, this.SharedStringsUri),
                                         TargetMode.Internal,
                                         ExcelPackage.schemaRelationships + "/sharedStrings");
        }

        part.SaveHandler = this.SaveSharedStringHandler;

        //// Data validation
        this.ValidateDataValidations();

        //VBA
        if (this._vba != null)
        {
            this.VbaProject.Save();
        }
    }

    private void SaveExternalLinks()
    {
        foreach (ExcelExternalLink er in this._externalLinks)
        {
            if (er.Part == null)
            {
                ExcelExternalWorkbook ewb = er.As.ExternalWorkbook;
                Uri uri = GetNewUri(this._package.ZipPackage, "/xl/externalLinks/externalLink{0}.xml");
                ewb.Part = this._package.ZipPackage.CreatePart(uri, ContentTypes.contentTypeExternalLink);
                FileInfo extFile = ((ExcelExternalWorkbook)er).File;
                ewb.Relation = er.Part.CreateRelationship(extFile.FullName, TargetMode.External, ExcelPackage.schemaRelationships + "/externalLinkPath");

                ZipPackageRelationship wbRel = this.Part.CreateRelationship(uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/externalLink");
                XmlElement wbExtRefElement = (XmlElement)this.CreateNode("d:externalReferences/d:externalReference", false, true);
                _ = wbExtRefElement.SetAttribute("id", ExcelPackage.schemaRelationships, wbRel.Id);
            }

            StreamWriter sw = new StreamWriter(er.Part.GetStream(FileMode.CreateNew));
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">");
            er.Save(sw);
            sw.Write("</externalLink>");
            sw.Flush();
        }
    }

    private void SavePivotTableCaches()
    {
        foreach (PivotTableCacheRangeInfo info in this._pivotTableCaches.Values)
        {
            foreach (PivotTableCacheInternal cache in info.PivotCaches)
            {
                if (cache._pivotTables.Count == 0)
                {
                    cache.Delete();

                    continue;
                }

                //Rewrite the pivottable address again if any rows or columns have been inserted or deleted
                ExcelRangeBase? r = cache.SourceRange;

                if (r != null && r.Worksheet != null) //Source does not exist
                {
                    ExcelTable t = ExcelTableCollection.GetFromRange(r);

                    XmlNodeList? fields = cache.CacheDefinitionXml.SelectNodes("d:pivotCacheDefinition/d:cacheFields/d:cacheField", this.NameSpaceManager);

                    if (fields != null)
                    {
                        this.FixFieldNamesAndUpdateSharedItems(cache, t, fields);
                    }

                    cache.RefreshOnLoad = true;
                    cache.CacheDefinitionXml.Save(cache.Part.GetStream(FileMode.Create));
                    cache.ResetRecordXml(this._package.ZipPackage);
                }
            }
        }
    }

    private void FixFieldNamesAndUpdateSharedItems(PivotTableCacheInternal cache, ExcelTable t, XmlNodeList fields)
    {
        cache.RefreshFields();
        int ix = 0;
        HashSet<string> flds = new HashSet<string>();
        ExcelRangeBase sourceRange = cache.SourceRange;

        foreach (XmlElement node in fields)
        {
            if (ix >= sourceRange.Columns)
            {
                break;
            }

            string? fldName = node.GetAttribute("name"); //Fixes issue 15295 dup name error

            if (string.IsNullOrEmpty(fldName))
            {
                fldName = t == null ? sourceRange.Offset(0, ix, 1, 1).Value.ToString() : t.Columns[ix].Name;
            }

            if (flds.Contains(fldName))
            {
                fldName = GetNewName(flds, fldName);
            }

            _ = flds.Add(fldName);
            node.SetAttribute("name", fldName);

            if (cache.Fields[ix].Grouping == null)
            {
                cache.Fields[ix].WriteSharedItems(node, this.NameSpaceManager);
            }

            ix++;
        }
    }

    private static string GetNewName(HashSet<string> flds, string fldName)
    {
        int ix = 2;

        while (flds.Contains(fldName + ix.ToString(CultureInfo.InvariantCulture)))
        {
            ix++;
        }

        return fldName + ix.ToString(CultureInfo.InvariantCulture);
    }

    private void DeleteCalcChain()
    {
        //Remove the calc chain if it exists.
        Uri uriCalcChain = new Uri("/xl/calcChain.xml", UriKind.Relative);

        if (this._package.ZipPackage.PartExists(uriCalcChain))
        {
            Uri calcChain = new Uri("calcChain.xml", UriKind.Relative);

            foreach (ZipPackageRelationship relationship in this._package.Workbook.Part.GetRelationships())
            {
                if (relationship.TargetUri == calcChain)
                {
                    this._package.Workbook.Part.DeleteRelationship(relationship.Id);

                    break;
                }
            }

            // delete the calcChain part
            this._package.ZipPackage.DeletePart(uriCalcChain);
        }
    }

    private void ValidateDataValidations()
    {
        foreach (ExcelWorksheet sheet in this._package.Workbook.Worksheets)
        {
            if (!(sheet is ExcelChartsheet) && sheet.DataValidations != null)
            {
                sheet.DataValidations.ValidateAll();
            }
        }
    }

    private void SaveSharedStringHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
    {
        //Init Zip
        stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
        _ = stream.PutNextEntry(fileName);

        StringBuilder cache = new StringBuilder();
        Encoding utf8Encoder = Encoding.GetEncoding("UTF-8", new EncoderReplacementFallback(string.Empty), new DecoderReplacementFallback(string.Empty));
        StreamWriter sw = new StreamWriter(stream, utf8Encoder);

        _ = cache.AppendFormat("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">",
                           this._sharedStrings.Count);

        foreach (string t in this._sharedStrings.Keys)
        {
            SharedStringItem ssi = this._sharedStrings[t];

            if (ssi.isRichText)
            {
                _ = cache.Append("<si>");
                ConvertUtil.ExcelEncodeString(cache, t);
                _ = cache.Append("</si>");
            }
            else
            {
                if (t.Length > 0
                    && (t[0] == ' '
                        || t[t.Length - 1] == ' '
                        || t.Contains("  ")
                        || t.Contains("\r")
                        || t.Contains("\t")
                        || t.Contains("\n"))) //Fixes issue 14849
                {
                    _ = cache.Append("<si><t xml:space=\"preserve\">");
                }
                else
                {
                    _ = cache.Append("<si><t>");
                }

                ConvertUtil.ExcelEncodeString(cache, ConvertUtil.ExcelEscapeString(t));
                _ = cache.Append("</t></si>");
            }

            if (cache.Length > 0x600000)
            {
                sw.Write(cache.ToString());
                cache = new StringBuilder();
            }
        }

        _ = cache.Append("</sst>");
        sw.Write(cache.ToString());
        sw.Flush();

        // Issue 15252: Save SharedStrings only once
        //Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, SharedStringsUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/sharedStrings");
    }

    private void UpdateDefinedNamesXml()
    {
        try
        {
            XmlNode top = this.WorkbookXml.SelectSingleNode("//d:definedNames", this.NameSpaceManager);

            if (!this.ExistsNames())
            {
                if (top != null)
                {
                    _ = this.TopNode.RemoveChild(top);
                }

                return;
            }
            else
            {
                if (top == null)
                {
                    _ = this.CreateNode("d:definedNames");
                    top = this.WorkbookXml.SelectSingleNode("//d:definedNames", this.NameSpaceManager);
                }
                else
                {
                    top.RemoveAll();
                }

                foreach (ExcelNamedRange name in this._names)
                {
                    XmlElement elem = this.WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
                    _ = top.AppendChild(elem);
                    elem.SetAttribute("name", name.Name);

                    if (name.IsNameHidden)
                    {
                        elem.SetAttribute("hidden", "1");
                    }

                    if (!string.IsNullOrEmpty(name.NameComment))
                    {
                        elem.SetAttribute("comment", name.NameComment);
                    }

                    SetNameElement(name, elem);
                }
            }

            foreach (ExcelWorksheet ws in this._worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    foreach (ExcelNamedRange name in ws.Names)
                    {
                        XmlElement elem = this.WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
                        _ = top.AppendChild(elem);
                        elem.SetAttribute("name", name.Name);
                        elem.SetAttribute("localSheetId", name.LocalSheetId.ToString());

                        if (name.IsNameHidden)
                        {
                            elem.SetAttribute("hidden", "1");
                        }

                        if (!string.IsNullOrEmpty(name.NameComment))
                        {
                            elem.SetAttribute("comment", name.NameComment);
                        }

                        SetNameElement(name, elem);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Internal error updating named ranges ", ex);
        }
    }

    private static void SetNameElement(ExcelNamedRange name, XmlElement elem)
    {
        if (name.IsName)
        {
            if (string.IsNullOrEmpty(name.NameFormula))
            {
                if (TypeCompat.IsPrimitive(name.NameValue) || name.NameValue is double || name.NameValue is decimal)
                {
                    elem.InnerText = Convert.ToDouble(name.NameValue, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
                }
                else if (name.NameValue is DateTime)
                {
                    elem.InnerText = ((DateTime)name.NameValue).ToOADate().ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    elem.InnerText = "\"" + name.NameValue.ToString() + "\"";
                }
            }
            else
            {
                elem.InnerText = name.NameFormula;
            }
        }
        else
        {
            if (name.Table != null)
            {
                elem.InnerText = name.Address;
            }
            else if (name.AllowRelativeAddress)
            {
                elem.InnerText = name.FullAddress;
            }
            else
            {
                elem.InnerText = name.FullAddressAbsolute;
            }
        }
    }

    /// <summary>
    /// Is their any names in the workbook or in the sheets.
    /// </summary>
    /// <returns>?</returns>
    private bool ExistsNames()
    {
        if (this._names.Count == 0)
        {
            foreach (ExcelWorksheet ws in this.Worksheets)
            {
                if (ws is ExcelChartsheet)
                {
                    continue;
                }

                if (ws.Names.Count > 0)
                {
                    return true;
                }
            }
        }
        else
        {
            return true;
        }

        return false;
    }

    #endregion

    #endregion

    /// <summary>
    /// Removes all formulas within the entire workbook, but keeps the calculated values.
    /// </summary>
    public void ClearFormulas()
    {
        if (this.Worksheets == null || this.Worksheets.Count == 0)
        {
            return;
        }

        foreach (ExcelWorksheet worksheet in this.Worksheets)
        {
            worksheet.ClearFormulas();
        }
    }

    /// <summary>
    /// Removes all values of cells with formulas in the entire workbook, but keeps the formulas.
    /// </summary>
    public void ClearFormulaValues()
    {
        if (this.Worksheets == null || this.Worksheets.Count == 0)
        {
            return;
        }

        foreach (ExcelWorksheet worksheet in this.Worksheets)
        {
            worksheet.ClearFormulaValues();
        }
    }

    internal bool ExistsTableName(string Name)
    {
        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (ws is ExcelChartsheet)
            {
                continue;
            }

            if (ws.Tables._tableNames.ContainsKey(Name))
            {
                return true;
            }
        }

        return false;
    }

    internal bool ExistsPivotTableName(string Name)
    {
        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (ws is ExcelChartsheet)
            {
                continue;
            }

            if (ws.PivotTables._pivotTableNames.ContainsKey(Name))
            {
                return true;
            }
        }

        return false;
    }

    internal void AddPivotTableCache(PivotTableCacheInternal cacheReference, bool createWorkbookElement = true)
    {
        if (createWorkbookElement)
        {
            _ = this.CreateNode("d:pivotCaches");

            XmlElement item = this.WorkbookXml.CreateElement("pivotCache", ExcelPackage.schemaMain);
            item.SetAttribute("cacheId", cacheReference.CacheId.ToString());

            ZipPackageRelationship rel = this.Part.CreateRelationship(UriHelper.ResolvePartUri(this.WorkbookUri, cacheReference.CacheDefinitionUri),
                                                                      TargetMode.Internal,
                                                                      ExcelPackage.schemaRelationships + "/pivotCacheDefinition");

            _ = item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

            XmlNode? pivotCaches = this.WorkbookXml.SelectSingleNode("//d:pivotCaches", this.NameSpaceManager);
            _ = pivotCaches.AppendChild(item);
        }

        if (cacheReference.CacheSource == eSourceType.Worksheet && cacheReference.SourceRange != null)
        {
            string address;

            if (string.IsNullOrEmpty(cacheReference.SourceName))
            {
                address = cacheReference.SourceRange.FullAddress;
            }
            else
            {
                address = cacheReference.SourceName;
            }

            if (this._pivotTableCaches.TryGetValue(address, out PivotTableCacheRangeInfo cacheInfo))
            {
                cacheInfo.PivotCaches.Add(cacheReference);
            }
            else
            {
                this._pivotTableCaches.Add(address,
                                           new PivotTableCacheRangeInfo()
                                           {
                                               Address = address, PivotCaches = new List<PivotTableCacheInternal>() { cacheReference }
                                           });
            }
        }
    }

    internal void RemovePivotTableCache(int cacheId)
    {
        string path = $"d:pivotCaches/d:pivotCache[@cacheId={cacheId}]";
        string relId = this.GetXmlNodeString(path + "/@r:id");
        this.DeleteNode(path, true);
        this.Part.DeleteRelationship(relId);
    }

    //internal bool _isCalculated=false;
    /// <summary>
    /// Disposes the workbooks
    /// </summary>
    public void Dispose()
    {
        if (this._sharedStrings != null)
        {
            this._sharedStrings.Clear();
            this._sharedStrings = null;
        }

        if (this._sharedStringsList != null)
        {
            this._sharedStringsList.Clear();
            this._sharedStringsList = null;
        }

        this._vba = null;

        if (this._worksheets != null)
        {
            this._worksheets.Dispose();
            this._worksheets = null;
        }

        this._package = null;
        this._properties = null;

        if (this._formulaParser != null)
        {
            this._formulaParser.Dispose();
            this._formulaParser = null;
        }
    }

    /// <summary>
    /// Returns true if the workbook has pivot tables in any worksheet.
    /// </summary>
    public bool HasLoadedPivotTables
    {
        get
        {
            if (this._worksheets == null)
            {
                return false;
            }

            foreach (ExcelWorksheet ws in this._worksheets)
            {
                if (ws.HasLoadedPivotTables == true)
                {
                    return true;
                }
            }

            return false;
        }
    }

    internal void ReadAllPivotTables()
    {
        if (this._nextPivotTableID > 0)
        {
            return;
        }

        this._nextPivotTableID = 1;

        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (!(ws is ExcelChartsheet)) //Chartsheets should be ignored.
            {
                foreach (ExcelPivotTable pt in ws.PivotTables)
                {
                    if (pt.CacheId >= this._nextPivotTableID)
                    {
                        this._nextPivotTableID = pt.CacheId + 1;
                    }
                }
            }
        }
    }

    internal void ReadAllTables()
    {
        if (this._nextTableID > 0)
        {
            return;
        }

        this._nextTableID = 1;

        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (!(ws is ExcelChartsheet)) //Chartsheets should be ignored.
            {
                foreach (ExcelTable tbl in ws.Tables)
                {
                    if (tbl.Id >= this._nextTableID)
                    {
                        this._nextTableID = tbl.Id + 1;
                    }
                }
            }
        }
    }

    internal Dictionary<string, ExcelSlicerCache> _slicerCaches;

    internal Dictionary<string, ExcelSlicerCache> SlicerCaches
    {
        get
        {
            if (this._slicerCaches == null)
            {
                this.LoadSlicerCaches();
            }

            return this._slicerCaches;
        }
    }

    internal ExcelSlicerCache GetSlicerCaches(string key)
    {
        if (this._slicerCaches == null)
        {
            this.LoadSlicerCaches();
        }

        if (this._slicerCaches != null && this._slicerCaches.TryGetValue(key, out ExcelSlicerCache c))
        {
            return c;
        }
        else
        {
            return null;
        }
    }

    internal void LoadSlicerCaches()
    {
        this._slicerCaches = new Dictionary<string, ExcelSlicerCache>();

        foreach (ZipPackageRelationship r in this.Part.GetRelationshipsByType(ExcelPackage.schemaRelationshipsSlicerCache))
        {
            Uri uri = UriHelper.ResolvePartUri(this.WorkbookUri, r.TargetUri);
            ZipPackagePart p = this.Part.Package.GetPart(uri);
            XmlDocument xml = new XmlDocument();
            LoadXmlSafe(xml, p.GetStream());

            ExcelSlicerCache cache;

            if (xml.DocumentElement.FirstChild.LocalName == "pivotTables")
            {
                cache = new ExcelPivotTableSlicerCache(this.NameSpaceManager);
            }
            else
            {
                cache = new ExcelTableSlicerCache(this.NameSpaceManager);
            }

            cache.Uri = uri;
            cache.CacheRel = r;
            cache.Part = p;
            cache.TopNode = xml.DocumentElement;
            cache.SlicerCacheXml = xml;
            cache.Init(this);

            this._slicerCaches.Add(cache.Name, cache);
        }
    }

    internal ExcelTable GetTable(int tableId)
    {
        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            ExcelTable? t = ws.Tables.FirstOrDefault(x => x.Id == tableId);

            if (t != null)
            {
                return t;
            }
        }

        return null;
    }

    internal void ClearDefaultHeightsAndWidths()
    {
        foreach (ExcelWorksheet ws in this.Worksheets)
        {
            if (ws.IsChartSheet == false)
            {
                if (ws.CustomHeight == false)
                {
                    ws._defaultRowHeight = double.NaN;
                }
            }
        }
    }
} // end Workbook