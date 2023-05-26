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
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Linq;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Drawing.Slicer;
using System.Text;
using System.Runtime.InteropServices.ComTypes;
using OfficeOpenXml.Constants;
using System.Xml.Linq;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml;

/// <summary>
/// The collection of worksheets for the workbook
/// </summary>
public class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IDisposable
{
    #region Private Properties
    internal ExcelPackage _pck;
    internal ChangeableDictionary<ExcelWorksheet> _worksheets;
    private XmlNamespaceManager _namespaceManager;
    #endregion
    #region ExcelWorksheets Constructor
    internal ExcelWorksheets(ExcelPackage pck, XmlNamespaceManager nsm, XmlNode topNode) :
        base(nsm, topNode)
    {
        this._pck = pck;
        this._namespaceManager = nsm;
        int ix = 0;
        this._worksheets = new ChangeableDictionary<ExcelWorksheet>();

        foreach (XmlNode sheetNode in topNode.ChildNodes)
        {
            if (sheetNode.NodeType == XmlNodeType.Element)
            {
                string name = sheetNode.Attributes["name"].Value;
                //Get the relationship id
                string relId = sheetNode.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships).Value;
                int sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);

                if (string.IsNullOrEmpty(relId))
                {
                    ExcelWorksheet? ws = this.AddSheet(name, false, null, null, (XmlElement)sheetNode);
                    ws.SheetId = sheetID;
                    //_worksheets.Add(ix, ws);
                }
                else
                {
                    ZipPackageRelationship? sheetRelation = pck.Workbook.Part.GetRelationship(relId);
                    Uri uriWorksheet = UriHelper.ResolvePartUri(pck.Workbook.WorkbookUri, sheetRelation.TargetUri);

                    int positionID = ix + this._pck._worksheetAdd;
                    //add the worksheet
                    if (sheetRelation.RelationshipType.EndsWith("chartsheet"))
                    {
                        this._worksheets.Add(ix, new ExcelChartsheet(this._namespaceManager, this._pck, relId, uriWorksheet, name, sheetID, positionID, null));
                    }
                    else
                    {
                        this._worksheets.Add(ix, new ExcelWorksheet(this._namespaceManager, this._pck, relId, uriWorksheet, name, sheetID, positionID, null));
                    }
                }
                ix++;
            }
        }
    }

    private static eWorkSheetHidden TranslateHidden(string value)
    {
        switch (value)
        {
            case "hidden":
                return eWorkSheetHidden.Hidden;
            case "veryHidden":
                return eWorkSheetHidden.VeryHidden;
            default:
                return eWorkSheetHidden.Visible;
        }
    }
    #endregion
    #region ExcelWorksheets Public Properties
    /// <summary>
    /// Returns the number of worksheets in the workbook
    /// </summary>
    public int Count
    {
        get { return (this._worksheets.Count); }
    }
    #endregion
    internal const string ERR_DUP_WORKSHEET = "A worksheet with this name already exists in the workbook";
    internal const string WORKSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    internal const string CHARTSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
    #region ExcelWorksheets Public Methods
    /// <summary>
    /// Foreach support
    /// </summary>
    /// <returns>An enumerator</returns>
    public IEnumerator<ExcelWorksheet> GetEnumerator()
    {
        return (this._worksheets.GetEnumerator());
    }
    #region IEnumerable Members

    IEnumerator IEnumerable.GetEnumerator()
    {
        return (this._worksheets.GetEnumerator());
    }

    #endregion
    #region Add Worksheet
    /// <summary>
    /// Adds a new blank worksheet.
    /// </summary>
    /// <param name="Name">The name of the workbook</param>
    public ExcelWorksheet Add(string Name)
    {
        ExcelWorksheet worksheet = this.AddSheet(Name, false, null);
        return worksheet;
    }
    private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType, ExcelPivotTable pivotTableSource = null, XmlElement sheetElement=null)
    {   
        lock (this._worksheets)
        {
            Name = ValidateFixSheetName(Name);
            if (this.GetByName(Name) != null)
            {
                throw (new InvalidOperationException(ERR_DUP_WORKSHEET + " : " + Name));
            }

            this.GetSheetURI(ref Name, out int sheetID, out Uri uriWorksheet, isChart);
            ZipPackagePart worksheetPart = this._pck.ZipPackage.CreatePart(uriWorksheet, isChart ? CHARTSHEET_CONTENTTYPE : WORKSHEET_CONTENTTYPE, this._pck.Compression);

            //Create the new, empty worksheet and save it to the package
            StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
            XmlDocument worksheetXml = CreateNewWorksheet(isChart);
            worksheetXml.Save(streamWorksheet);
            ZipPackage.Flush();

            string rel = this.CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart, sheetElement);

            int positionID = this._worksheets.Count + this._pck._worksheetAdd;
            ExcelWorksheet worksheet;
            if (isChart)
            {
                worksheet = new ExcelChartsheet(this._namespaceManager, this._pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible, (eChartType)chartType, pivotTableSource);
            }
            else
            {
                worksheet = new ExcelWorksheet(this._namespaceManager, this._pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);
            }

            this._worksheets.Add(this._worksheets.Count, worksheet);
            if (this._pck.Workbook.VbaProject != null)
            {
                string? name = this._pck.Workbook.VbaProject.GetModuleNameFromWorksheet(worksheet);
                this._pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(worksheet.CodeNameChange) { Name = name, Code = "", Attributes = ExcelVbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                worksheet.CodeModuleName = name;
            }

            return worksheet;
        }
    }
    /// <summary>
    /// Adds a copy of a worksheet
    /// </summary>
    /// <param name="Name">The name of the workbook</param>
    /// <param name="Copy">The worksheet to be copied</param>
    public ExcelWorksheet Add(string Name, ExcelWorksheet Copy)
    {
        lock (this._worksheets)
        {
            return WorksheetCopyHelper.Copy(this, Name, Copy);
        }
    }
    /// <summary>
    /// Adds a chartsheet to the workbook.
    /// </summary>
    /// <param name="Name">The name of the worksheet</param>
    /// <param name="chartType">The type of chart</param>
    /// <returns></returns>
    public ExcelChartsheet AddChart(string Name, eChartType chartType)
    {
        if (ExcelChart.IsTypeStock(chartType))
        {
            throw (new InvalidOperationException("Please use method AddStockChart for Stock Charts"));
        }
        return (ExcelChartsheet)this.AddSheet(Name, true, chartType, null);
    }
    /// <summary>
    /// Adds a chartsheet to the workbook.
    /// </summary>
    /// <param name="Name">The name of the worksheet</param>
    /// <param name="chartType">The type of chart</param>
    /// <param name="pivotTableSource">The pivottable source</param>
    /// <returns></returns>
    public ExcelChartsheet AddChart(string Name, eChartType chartType, ExcelPivotTable pivotTableSource)
    {
        return (ExcelChartsheet)this.AddSheet(Name, true, chartType, pivotTableSource);
    }
    /// <summary>
    /// Adds a stock chart sheet to the workbook.
    /// </summary>
    /// <param name="Name">The name of the worksheet</param>
    /// <param name="CategorySerie">The category serie. A serie containing dates or names</param>
    /// <param name="HighSerie">The high price serie</param>    
    /// <param name="LowSerie">The low price serie</param>    
    /// <param name="CloseSerie">The close price serie containing</param>    
    /// <param name="OpenSerie">The opening price serie. Supplying this serie will create a StockOHLC or StockVOHLC chart</param>
    /// <param name="VolumeSerie">The volume represented as a column chart. Supplying this serie will create a StockVHLC or StockVOHLC chart</param>
    /// <returns></returns>
    public ExcelChartsheet AddStockChart(string Name, ExcelRangeBase CategorySerie, ExcelRangeBase HighSerie, ExcelRangeBase LowSerie, ExcelRangeBase CloseSerie, ExcelRangeBase OpenSerie = null, ExcelRangeBase VolumeSerie = null)
    {
        eChartType chartType = ExcelStockChart.GetChartType(OpenSerie, VolumeSerie);
        ExcelChartsheet? sheet = (ExcelChartsheet)this.AddSheet(Name, true, chartType, null);
        ExcelStockChart? chart = (ExcelStockChart)sheet.Chart;
        ExcelStockChart.SetStockChartSeries(chart, chartType, CategorySerie.FullAddress, HighSerie.FullAddress, LowSerie.FullAddress, CloseSerie.FullAddress, OpenSerie?.FullAddress, VolumeSerie?.FullAddress);
        return sheet;
    }
    internal int? GetFirstVisibleSheetIndex()
    {
        for (int i = 0; i < this._worksheets.Count; i++)
        {
            if (this._worksheets[i].Hidden == eWorkSheetHidden.Visible)
            {
                return i;
            }
        }
        throw new InvalidOperationException("The worksheets collection must have at least one visible woreksheet");
    }


    internal string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet, bool isChart, XmlElement sheetElement)
    {
        //Create the relationship between the workbook and the new worksheet
        ZipPackageRelationship? rel = this._pck.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(this._pck.Workbook.WorkbookUri, uriWorksheet), TargetMode.Internal, ExcelPackage.schemaRelationships + "/" + (isChart ? "chartsheet" : "worksheet"));
        ZipPackage.Flush();

        //Create the new sheet node
        if(sheetElement==null)
        {
            sheetElement = this._pck.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
            sheetElement.SetAttribute("name", Name);
            sheetElement.SetAttribute("sheetId", sheetID.ToString());
            this.TopNode.AppendChild(sheetElement);
        }
        sheetElement.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

        return rel.Id;
    }

    internal void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet, bool isChart)
    {
        Name = ValidateFixSheetName(Name);
        sheetID = this.Any() ? this.Max(ws => ws.SheetId) + 1 : 1;
        int uriId = sheetID;


        // get the next available worhsheet uri
        do
        {
            if (isChart)
            {
                uriWorksheet = new Uri("/xl/chartsheets/chartsheet" + uriId + ".xml", UriKind.Relative);
            }
            else
            {
                uriWorksheet = new Uri("/xl/worksheets/sheet" + uriId + ".xml", UriKind.Relative);
            }

            uriId++;
        } while (this._pck.ZipPackage.PartExists(uriWorksheet));
    }

    internal static string ValidateFixSheetName(string Name)
    {
        if (string.IsNullOrEmpty(Name) || Name.Trim() == "")
        {
            throw new ArgumentException("The worksheet cannot have an empty name");
        }

        //remove invalid characters
        if (ValidateName(Name))
        {
            if (Name.IndexOf(':') > -1)
            {
                Name = Name.Replace(':', ' ');
            }

            if (Name.IndexOf('/') > -1)
            {
                Name = Name.Replace('/', ' ');
            }

            if (Name.IndexOf('\\') > -1)
            {
                Name = Name.Replace('\\', ' ');
            }

            if (Name.IndexOf('?') > -1)
            {
                Name = Name.Replace('?', ' ');
            }

            if (Name.IndexOf('[') > -1)
            {
                Name = Name.Replace('[', ' ');
            }

            if (Name.IndexOf(']') > -1)
            {
                Name = Name.Replace(']', ' ');
            }
        }

        if (Name.StartsWith("'", StringComparison.OrdinalIgnoreCase) || Name.EndsWith("'", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The worksheet name cannot start or end with an apostrophe (').", "Name");
        }
        if (Name.Length > 31)
        {
            Name = Name.Substring(0, 31);   //A sheet can have max 31 char's            
        }

        return Name;
    }
    /// <summary>
    /// Validate the sheetname
    /// </summary>
    /// <param name="Name">The Name</param>
    /// <returns>True if valid</returns>
    private static bool ValidateName(string Name)
    {
        return System.Text.RegularExpressions.Regex.IsMatch(Name, @":|\?|/|\\|\[|\]");
    }

    /// <summary>
    /// Creates the XML document representing a new empty worksheet
    /// </summary>
    /// <returns></returns>
    internal static XmlDocument CreateNewWorksheet(bool isChart)
    {
        XmlDocument xmlDoc = new XmlDocument();
        XmlElement elemWs = xmlDoc.CreateElement(isChart ? "chartsheet" : "worksheet", ExcelPackage.schemaMain);
        elemWs.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
        xmlDoc.AppendChild(elemWs);


        if (isChart)
        {
            XmlElement elemSheetPr = xmlDoc.CreateElement("sheetPr", ExcelPackage.schemaMain);
            elemWs.AppendChild(elemSheetPr);

            XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
            elemWs.AppendChild(elemSheetViews);

            XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
            elemSheetView.SetAttribute("workbookViewId", "0");
            elemSheetView.SetAttribute("zoomToFit", "1");

            elemSheetViews.AppendChild(elemSheetView);
        }
        else
        {
            XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
            elemWs.AppendChild(elemSheetViews);

            XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
            elemSheetView.SetAttribute("workbookViewId", "0");
            elemSheetViews.AppendChild(elemSheetView);

            XmlElement elemSheetFormatPr = xmlDoc.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
            elemSheetFormatPr.SetAttribute("defaultRowHeight", "15");
            elemWs.AppendChild(elemSheetFormatPr);

            XmlElement elemSheetData = xmlDoc.CreateElement("sheetData", ExcelPackage.schemaMain);
            elemWs.AppendChild(elemSheetData);
        }
        return xmlDoc;
    }
    #endregion
    #region Delete Worksheet
    /// <summary>
    /// Deletes a worksheet from the collection
    /// </summary>
    /// <param name="Index">The position of the worksheet in the workbook</param>
    public void Delete(int Index)
    {            
        /*
        * Hack to prefetch all the drawings,
        * so that all the images are referenced, 
        * to prevent the deletion of the image file, 
        * when referenced more than once
        */
        foreach (ExcelWorksheet? ws in this._worksheets)
        {
            ExcelDrawings? drawings = ws.Drawings;
        }

        ExcelWorksheet worksheet = this._worksheets[Index - this._pck._worksheetAdd];
        if (worksheet.Drawings.Count > 0)
        {
            worksheet.Drawings.ClearDrawings();
        }

        //Remove all comments
        if (!(worksheet is ExcelChartsheet) && worksheet.Comments.Count > 0)
        {
            worksheet.Comments.Clear();
        }

        while(worksheet.PivotTables.Count>0)
        {
            worksheet.PivotTables.Delete(worksheet.PivotTables[0]);
        }
        //Delete any parts still with relations to the Worksheet.
        this.DeleteRelationsAndParts(worksheet.Part);


        //Delete the worksheet part and relation from the package 
        this._pck.Workbook.Part.DeleteRelationship(worksheet.RelationshipId);

        //Delete worksheet from the workbook XML
        XmlNode sheetsNode = this._pck.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", this._namespaceManager);
        if (sheetsNode != null)
        {
            XmlNode sheetNode = sheetsNode.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetId), this._namespaceManager);
            if (sheetNode != null)
            {
                sheetsNode.RemoveChild(sheetNode);
            }
        }
        if (this._pck.Workbook.VbaProject != null)
        {
            this._pck.Workbook.VbaProject.Modules.Remove(worksheet.CodeModule);
        }

        this._worksheets.RemoveAndShift(Index - this._pck._worksheetAdd);
        this.ReindexWorksheetDictionary();
        //If the active sheet is deleted, set the first tab as active.
        if (this._pck.Workbook.Worksheets.Count > 0)
        {
            if (this._pck.Workbook.View.ActiveTab >= this._pck.Workbook.Worksheets.Count)
            {
                this._pck.Workbook.View.ActiveTab = Math.Min(this._pck.Workbook.View.ActiveTab - 1, this._pck.Workbook.Worksheets.Count-1);
            }
            if (this._pck.Workbook.View.ActiveTab == worksheet.SheetId)
            {
                this._pck.Workbook.Worksheets[this._pck._worksheetAdd].View.TabSelected = true;
            }
        }
    }

    private void DeleteRelationsAndParts(ZipPackagePart part)
    {
        List<ZipPackageRelationship>? rels = part.GetRelationships().ToList();
        for (int i = 0; i < rels.Count; i++)
        {
            ZipPackageRelationship? rel = rels[i];
            if (rel.RelationshipType != ExcelPackage.schemaImage && rel.TargetMode == TargetMode.Internal)
            {
                Uri? relUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                if (this._pck.ZipPackage.PartExists(relUri))
                {
                    this.DeleteRelationsAndParts(this._pck.ZipPackage.GetPart(relUri));
                }
            }
            part.DeleteRelationship(rel.Id);
        }

        this._pck.ZipPackage.DeletePart(part.Uri);
    }

    /// <summary>
    /// Deletes a worksheet from the collection
    /// </summary>
    /// <param name="name">The name of the worksheet in the workbook</param>
    public void Delete(string name)
    {
        ExcelWorksheet? sheet = this[name];
        if (sheet == null)
        {
            throw new ArgumentException(string.Format("Could not find worksheet to delete '{0}'", name));
        }

        this.Delete(sheet.PositionId);
    }
    /// <summary>
    /// Delete a worksheet from the collection
    /// </summary>
    /// <param name="Worksheet">The worksheet to delete</param>
    public void Delete(ExcelWorksheet Worksheet)
    {
        int ix = Worksheet.PositionId - this._pck._worksheetAdd;
        if (ix < this._worksheets.Count && Worksheet == this._worksheets[ix])
        {
            this.Delete(Worksheet.PositionId);
        }
        else
        {
            throw (new ArgumentException("Worksheet is not in the collection."));
        }
    }
    #endregion
    internal void ReindexWorksheetDictionary()
    {
        int index = 0;
        ChangeableDictionary<ExcelWorksheet>? worksheets = new ChangeableDictionary<ExcelWorksheet>();
        foreach (ExcelWorksheet? entry in this._worksheets)
        {
            entry.PositionId = index + this._pck._worksheetAdd;
            worksheets.Add(index++, entry);
        }

        this._worksheets = worksheets;
    }

#if Core
    /// <summary>
    /// Returns the worksheet at the specified position. 
    /// </summary>
    /// <param name="PositionID">The position of the worksheet. Collection is zero-based or one-base depending on the Package.Compatibility.IsWorksheets1Based propery. Default is Zero based</param>
    /// <seealso cref="ExcelPackage.Compatibility"/>
    /// <returns></returns>
#else
        /// <summary>
        /// Returns the worksheet at the specified position. 
        /// </summary>
        /// <param name="PositionID">The position of the worksheet. Collection is zero-based or one-base depending on the Package.Compatibility.IsWorksheets1Based propery. Default is One based</param>
        /// <seealso cref="ExcelPackage.Compatibility"/>
        /// <returns></returns>
#endif
    public ExcelWorksheet this[int PositionID]
    {
        get
        {
            int ix = PositionID - this._pck._worksheetAdd;
            if (this._worksheets.ContainsKey(ix))
            {
                return this._worksheets[ix];
            }
            else
            {
                throw (new IndexOutOfRangeException("Worksheet position out of range."));
            }
        }
    }

    /// <summary>
    /// Returns the worksheet matching the specified name
    /// </summary>
    /// <param name="Name">The name of the worksheet</param>
    /// <returns></returns>
    public ExcelWorksheet this[string Name]
    {
        get
        {
            return this.GetByName(Name);
        }
    }
    /// <summary>
    /// Copies the named worksheet and creates a new worksheet in the same workbook
    /// </summary>
    /// <param name="Name">The name of the existing worksheet</param>
    /// <param name="NewName">The name of the new worksheet to create</param>
    /// <returns>The new copy added to the end of the worksheets collection</returns>
    public ExcelWorksheet Copy(string Name, string NewName)
    {
        ExcelWorksheet Copy = this[Name];
        if (Copy == null)
        {
            throw new ArgumentException(string.Format("Copy worksheet error: Could not find worksheet to copy '{0}'", Name));
        }

        ExcelWorksheet added = this.Add(NewName, Copy);
        return added;
    }
    #endregion
    internal ExcelWorksheet GetBySheetID(int localSheetID)
    {
        foreach (ExcelWorksheet ws in this)
        {
            if (ws.SheetId == localSheetID)
            {
                return ws;
            }
        }
        return null;
    }
    internal ExcelWorksheet GetByName(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return null;
        }

        name = ValidateFixSheetName(name);
        ExcelWorksheet ws = null;
        foreach (ExcelWorksheet worksheet in this._worksheets)
        {
            if (worksheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                ws = worksheet;
            }
        }
        return (ws);
    }

    /// <summary>
    /// Return a worksheet by its name. Can throw an exception if the worksheet does not exist.
    /// </summary>
    /// <param name="worksheetName">Name of the reqested worksheet</param>
    /// <param name="paramName">Name of the parameter</param>
    /// <param name="throwIfNull">Throws an <see cref="ArgumentNullException"></see> if the worksheet doesn't exist.</param>
    /// <returns></returns>
    private ExcelWorksheet GetWorksheetByName(string worksheetName, string paramName = null, bool throwIfNull = true)
    {
        ExcelWorksheet? worksheet = this.GetByName(worksheetName);
        if (worksheet == null && throwIfNull)
        {
            throw new ArgumentNullException(paramName ?? "worksheet", $"Could not find worksheet to move sourceName");
        }
        return worksheet;
    }
    internal bool _areDrawingsLoaded = false;
    //#region Move worksheet functions
    /// <summary>
    /// Moves the source worksheet to the position before the target worksheet
    /// </summary>
    /// <param name="sourceName">The name of the source worksheet</param>
    /// <param name="targetName">The name of the target worksheet</param>
    public void MoveBefore(string sourceName, string targetName)
    {
        MoveSheetXmlNode.RearrangeWorksheets(this, sourceName, targetName, true);
    }

    /// <summary>
    /// Moves the source worksheet to the position before the target worksheet
    /// </summary>
    /// <param name="sourcePositionId">The id of the source worksheet</param>
    /// <param name="targetPositionId">The id of the target worksheet</param>
    public void MoveBefore(int sourcePositionId, int targetPositionId)
    {
        MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, targetPositionId, true);
    }

    /// <summary>
    /// Moves the source worksheet to the position after the target worksheet
    /// </summary>
    /// <param name="sourceName">The name of the source worksheet</param>
    /// <param name="targetName">The name of the target worksheet</param>
    public void MoveAfter(string sourceName, string targetName)
    {
        MoveSheetXmlNode.RearrangeWorksheets(this, sourceName, targetName, false);
    }

    /// <summary>
    /// Moves the source worksheet to the position after the target worksheet
    /// </summary>
    /// <param name="sourcePositionId">The id of the source worksheet</param>
    /// <param name="targetPositionId">The id of the target worksheet</param>
    public void MoveAfter(int sourcePositionId, int targetPositionId)
    {
        MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, targetPositionId, true);
    }

    /// <summary>
    /// Moves the source worksheet to the start of the worksheets collection
    /// </summary>
    /// <param name="sourceName">The name of the source worksheet</param>
    public void MoveToStart(string sourceName)
    {
        Require.Argument(sourceName).IsNotNullOrEmpty("sourceName");
        ExcelWorksheet? worksheet = this.GetWorksheetByName(sourceName, "sourceName");
        this.MoveToStart(worksheet.PositionId);
    }
    /// <summary>
    /// Moves the source worksheet to the start of the worksheets collection
    /// </summary>
    /// <param name="sourcePositionId">The position of the source worksheet</param>
    public void MoveToStart(int sourcePositionId)
    {
        MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, this._pck._worksheetAdd, true);
    }

    /// <summary>
    /// Moves the source worksheet to the end of the worksheets collection
    /// </summary>
    /// <param name="sourceName">The name of the source worksheet</param>
    public void MoveToEnd(string sourceName)
    {
        Require.Argument(sourceName).IsNotNullOrEmpty("sourceName");
        ExcelWorksheet? worksheet = this.GetWorksheetByName(sourceName, "sourceName");
        this.MoveToEnd(worksheet.PositionId);
    }

    /// <summary>
    /// Moves the source worksheet to the end of the worksheets collection
    /// </summary>
    /// <param name="sourcePositionId">The position of the source worksheet</param>
    public void MoveToEnd(int sourcePositionId)
    {
        MoveSheetXmlNode.RearrangeWorksheets(this, sourcePositionId, this.Count - 1 + this._pck._worksheetAdd, false);
    }

    /// <summary>
    /// Dispose the worksheets collection
    /// </summary>
    public void Dispose()
    {
        if (this._worksheets != null)
        {
            foreach (ExcelWorksheet? sheet in this._worksheets)
            {
                ((IDisposable)sheet).Dispose();
            }

            this._worksheets = null;
            this._pck = null;
        }
    }

    internal void NormalStyleChange()
    {
        throw new NotImplementedException();
    }
} // end class Worksheets