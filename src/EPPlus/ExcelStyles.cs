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
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using draw = System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Slicer.Style;
using OfficeOpenXml.Table;
using System.Globalization;

namespace OfficeOpenXml;

/// <summary>
/// Containts all shared cell styles for a workbook
/// </summary>
public sealed class ExcelStyles : XmlHelper
{
    const string ColorsPath = "d:colors/d:indexedColors/d:rgbColor";
    const string NumberFormatsPath = "d:numFmts";
    const string FontsPath = "d:fonts";
    const string FillsPath = "d:fills";
    const string BordersPath = "d:borders";
    const string CellStyleXfsPath = "d:cellStyleXfs";
    const string CellXfsPath = "d:cellXfs";
    const string CellStylesPath = "d:cellStyles";
    const string TableStylesPath = "d:tableStyles";
    internal const string DxfsPath = "d:dxfs";
    internal const string DxfSlicerStylesPath = "d:extLst/d:ext[@uri='" + ExtLstUris.SlicerStylesDxfCollectionUri + "']/x14:dxfs";
    const string SlicerStylesPath = "d:extLst/d:ext[@uri='" + ExtLstUris.SlicerStylesUri + "']/x14:slicerStyles";
    XmlDocument _styleXml;
    internal ExcelWorkbook _wb;
    ExcelNamedStyleXml _normalStyle;
    XmlNamespaceManager _nameSpaceManager;
    internal int _nextDfxNumFmtID = 164;

    internal ExcelStyles(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb)
        : base(NameSpaceManager, xml.DocumentElement)
    {
        this._styleXml = xml;
        this._wb = wb;
        this._nameSpaceManager = NameSpaceManager;
        this.SchemaNodeOrder = new string[] { "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs" };
        this.LoadFromDocument();
    }

    /// <summary>
    /// Loads the style XML to memory
    /// </summary>
    private void LoadFromDocument()
    {
        //colors
        XmlNodeList colorNodes = this.GetNodes(ColorsPath);

        if (colorNodes != null && colorNodes.Count > 0)
        {
            int index = 0;

            foreach (XmlNode node in colorNodes)
            {
                ExcelColor.indexedColors[index++] = "#" + node.Attributes["rgb"].InnerText;
            }
        }

        //NumberFormats
        ExcelNumberFormatXml.AddBuildIn(this.NameSpaceManager, this.NumberFormats);
        XmlNode numNode = this.GetNode(NumberFormatsPath);

        if (numNode != null)
        {
            foreach (XmlNode n in numNode)
            {
                ExcelNumberFormatXml nf = new ExcelNumberFormatXml(this._nameSpaceManager, n);
                this.NumberFormats.Add(nf.Id, nf);

                if (nf.NumFmtId >= this.NumberFormats.NextId)
                {
                    this.NumberFormats.NextId = nf.NumFmtId + 1;
                }
            }
        }

        //Fonts
        XmlNode fontNode = this.GetNode(FontsPath);

        foreach (XmlNode n in fontNode)
        {
            ExcelFontXml f = new ExcelFontXml(this._nameSpaceManager, n);
            this.Fonts.Add(f.Id, f);
        }

        //Fills
        XmlNode fillNode = this.GetNode(FillsPath);

        foreach (XmlNode n in fillNode)
        {
            ExcelFillXml f;

            if (n.FirstChild != null && n.FirstChild.LocalName == "gradientFill")
            {
                f = new ExcelGradientFillXml(this._nameSpaceManager, n);
            }
            else
            {
                f = new ExcelFillXml(this._nameSpaceManager, n);
            }

            this.Fills.Add(f.Id, f);
        }

        //Borders
        XmlNode borderNode = this.GetNode(BordersPath);

        foreach (XmlNode n in borderNode)
        {
            ExcelBorderXml b = new ExcelBorderXml(this._nameSpaceManager, n);
            this.Borders.Add(b.Id, b);
        }

        //cellStyleXfs
        XmlNode styleXfsNode = this.GetNode(CellStyleXfsPath);

        if (styleXfsNode != null)
        {
            foreach (XmlNode n in styleXfsNode)
            {
                ExcelXfs item = new ExcelXfs(this._nameSpaceManager, n, this);
                this.CellStyleXfs.Add(item.Id, item);
            }
        }

        XmlNode styleNode = this.GetNode(CellXfsPath);

        for (int i = 0; i < styleNode.ChildNodes.Count; i++)
        {
            XmlNode n = styleNode.ChildNodes[i];
            ExcelXfs item = new ExcelXfs(this._nameSpaceManager, n, this);
            this.CellXfs.Add(item.Id, item);
        }

        //cellStyle
        XmlNode namedStyleNode = this.GetNode(CellStylesPath);

        if (namedStyleNode != null)
        {
            foreach (XmlNode n in namedStyleNode)
            {
                ExcelNamedStyleXml item = new ExcelNamedStyleXml(this._nameSpaceManager, n, this);

                if (item.BuildInId == 0)
                {
                    this._normalStyle = item;
                }

                this.NamedStyles.Add(item.Name, item);
            }
        }

        DxfStyleHandler.Load(this._wb, this, this.Dxfs, DxfsPath);
        this.LoadTableStyles();
        this.LoadSlicerStyles();
    }

    private void LoadSlicerStyles()
    {
        //Slicer Styles
        XmlNode slicerStylesNode = this.GetNode(SlicerStylesPath);

        if (slicerStylesNode != null)
        {
            DxfStyleHandler.Load(this._wb, this, this.DxfsSlicers, DxfSlicerStylesPath); //Slicer styles have their own dxf collection inside the extLst.

            foreach (XmlNode n in slicerStylesNode)
            {
                string? name = n.Attributes["name"]?.Value;
                XmlNode tableStyleNode;

                if (this._slicerTableStyleNodes.ContainsKey(name))
                {
                    tableStyleNode = this._slicerTableStyleNodes[name];
                }
                else if (this.TableStyles._dic.ContainsKey(name))
                {
                    tableStyleNode = this.TableStyles[name].TopNode;
                }
                else
                {
                    tableStyleNode = null;
                }

                ExcelSlicerNamedStyle? item = new ExcelSlicerNamedStyle(this._nameSpaceManager, n, tableStyleNode, this);
                this.SlicerStyles.Add(item.Name, item);
            }
        }
    }

    private void LoadTableStyles()
    {
        //Table Styles
        XmlNode tableStyleNode = this.GetNode(TableStylesPath);

        if (tableStyleNode != null)
        {
            foreach (XmlNode n in tableStyleNode)
            {
                bool pivot = !(n.Attributes["pivot"]?.Value == "0");
                bool table = !(n.Attributes["table"]?.Value == "0");

                if (pivot || table)
                {
                    ExcelTableNamedStyleBase item;

                    if (pivot == false)
                    {
                        item = new ExcelTableNamedStyle(this._nameSpaceManager, n, this);
                    }
                    else if (table == false)
                    {
                        item = new ExcelPivotTableNamedStyle(this._nameSpaceManager, n, this);
                    }
                    else
                    {
                        item = new ExcelTableAndPivotTableNamedStyle(this._nameSpaceManager, n, this);
                    }

                    this.TableStyles.Add(item.Name, item);
                }
                else
                {
                    //Styles for slicers and timelines. Timelines are currently unsupported.
                    string? name = n.Attributes["name"]?.Value;

                    if (string.IsNullOrEmpty(name) == false)
                    {
                        this._slicerTableStyleNodes.Add(name, n);
                    }
                }
            }
        }
    }

    internal ExcelNamedStyleXml GetNormalStyle()
    {
        if (this._normalStyle == null)
        {
            foreach (ExcelNamedStyleXml? style in this.NamedStyles)
            {
                if (style.BuildInId == 0)
                {
                    this._normalStyle = style;

                    break;
                }
            }

            if (this._normalStyle == null && this._wb.Styles.NamedStyles.Count > 0)
            {
                return this._wb.Styles.NamedStyles[0];
            }
        }

        return this._normalStyle;
    }

    internal ExcelStyle GetStyleObject(int Id, int PositionID, string Address)
    {
        if (Id < 0)
        {
            Id = 0;
        }

        return new ExcelStyle(this, this.PropertyChange, PositionID, Address, Id);
    }

    /// <summary>
    /// Handels changes of properties on the style objects
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    /// <returns></returns>
    internal int PropertyChange(StyleBase sender, StyleChangeEventArgs e)
    {
        ExcelAddressBase? address = new ExcelAddressBase(e.Address);
        ExcelWorksheet? ws = this._wb.Worksheets[e.PositionID];
        Dictionary<int, int> styleCashe = new Dictionary<int, int>();

        //Set single address
        lock (ws._values)
        {
            if (address.Addresses == null)
            {
                this.SetStyleAddress(sender, e, address, ws, ref styleCashe);
            }
            else
            {
                //Handle multiaddresses
                foreach (ExcelAddressBase? innerAddress in address.Addresses)
                {
                    this.SetStyleAddress(sender, e, innerAddress, ws, ref styleCashe);
                }
            }
        }

        return 0;
    }

    private void SetStyleAddress(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, ref Dictionary<int, int> styleCashe)
    {
        if (address.Start.Column == 0 || address.Start.Row == 0)
        {
            throw new Exception("error address");
        }

        //Columns
        else if (address.Start.Row == 1 && address.End.Row == ExcelPackage.MaxRows)
        {
            this.SetStyleFullColumn(sender, e, address, ws, styleCashe);
        }

        //Rows
        else if (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns)
        {
            this.SetStyleFullRow(sender, e, address, ws, styleCashe);
        }

        //Cellrange
        else
        {
            this.SetStyleCells(sender, e, address, ws, styleCashe);
        }
    }

    private void SetStyleCells(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
    {
        ws._values.EnsureColumnsExists(address._fromCol, address._toCol);
        Dictionary<int, int>? rowCache = new Dictionary<int, int>(address.End.Row - address.Start.Row + 1);
        Dictionary<int, ExcelValue>? colCache = new Dictionary<int, ExcelValue>(address.End.Column - address.Start.Column + 1);

        CellStoreEnumerator<ExcelValue>? cellEnum =
            new CellStoreEnumerator<ExcelValue>(ws._values, address.Start.Row, address.Start.Column, address.End.Row, address.End.Column);

        bool hasEnumValue = cellEnum.Next();

        for (int row = address._fromRow; row <= address._toRow; row++)
        {
            for (int col = address._fromCol; col <= address._toCol; col++)
            {
                ExcelValue value;

                if (hasEnumValue && row == cellEnum.Row && col == cellEnum.Column)
                {
                    value = cellEnum.Value;
                    hasEnumValue = cellEnum.Next();
                }
                else
                {
                    value = new ExcelValue { _styleId = 0 };
                }

                int s = value._styleId;

                if (s == 0)
                {
                    // get row styleId with cache
                    if (rowCache.ContainsKey(row))
                    {
                        s = rowCache[row];
                    }
                    else
                    {
                        s = ws._values.GetValue(row, 0)._styleId;
                        rowCache.Add(row, s);
                    }

                    if (s == 0)
                    {
                        // get column styleId with cache
                        if (colCache.ContainsKey(col))
                        {
                            s = colCache[col]._styleId;
                        }
                        else
                        {
                            ExcelValue v = ws._values.GetValue(0, col);

                            if (v._value == null)
                            {
                                if (colCache.TryGetValue(col, out ExcelValue ev))
                                {
                                    s = ev._styleId;
                                }
                                else
                                {
                                    int r = 0,
                                        c = col;

                                    if (ws._values.PrevCell(ref r, ref c))
                                    {
                                        if (!colCache.ContainsKey(c))
                                        {
                                            colCache.Add(c, ws._values.GetValue(0, c));
                                        }

                                        ExcelValue val = colCache[c];
                                        ExcelColumn? colObj = val._value as ExcelColumn;

                                        if (colObj != null && colObj.ColumnMax >= col) //Fixes issue 15174
                                        {
                                            s = val._styleId;
                                        }
                                    }
                                    else
                                    {
                                        colCache.Add(col, new ExcelValue() { _styleId = 0 });
                                    }
                                }
                            }
                            else
                            {
                                colCache.Add(col, v);
                                s = v._styleId;
                            }
                        }
                    }
                }

                if (styleCashe.ContainsKey(s))
                {
                    ws._values.SetValue(row, col, new ExcelValue { _value = value._value, _styleId = styleCashe[s] });
                }
                else
                {
                    ExcelXfs st;

                    if (s == 0)
                    {
                        ExcelNamedStyleXml? ns = this.GetNormalStyle(); //Get the xfs id for the normal style.

                        if (ns == null || ns.StyleXfId < 0)
                        {
                            st = this.CellXfs[0];
                        }
                        else
                        {
                            st = this.CellStyleXfs[ns.StyleXfId];
                        }
                    }
                    else
                    {
                        st = this.CellXfs[s];
                    }

                    int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                    styleCashe.Add(s, newId);
                    ws._values.SetValue(row, col, new ExcelValue { _value = value._value, _styleId = newId });
                }
            }
        }
    }

    private static bool GetFromCache(Dictionary<int, ExcelValue> colCache, int col, ref int s)
    {
        int c = col;

        while (!colCache.ContainsKey(--c))
        {
            if (c <= 0)
            {
                return false;
            }
        }

        ExcelColumn? colObj = (ExcelColumn)colCache[c]._value;

        if (colObj != null && colObj.ColumnMax >= col) //Fixes issue 15174
        {
            s = colCache[c]._styleId;
        }
        else
        {
            s = 0;
        }

        return true;
    }

    private void SetStyleFullRow(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
    {
        for (int rowNum = address.Start.Row; rowNum <= address.End.Row; rowNum++)
        {
            int s = ws.GetStyleInner(rowNum, 0);

            if (s == 0)
            {
                //iterate all columns and set the row to the style of the last column
                CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
                int cs = 0;

                while (cse.Next())
                {
                    cs = cse.Value._styleId;

                    if (cs == 0)
                    {
                        continue;
                    }

                    ExcelColumn? c = ws.GetValueInner(cse.Row, cse.Column) as ExcelColumn;

                    if (c != null && c.ColumnMax < ExcelPackage.MaxColumns)
                    {
                        for (int col = c.ColumnMin; col < c.ColumnMax; col++)
                        {
                            if (!ws.ExistsStyleInner(rowNum, col))
                            {
                                ws.SetStyleInner(rowNum, col, cs);
                            }
                        }
                    }
                }

                ws.SetStyleInner(rowNum, 0, cs);
                cse.Dispose();
            }

            if (styleCashe.ContainsKey(s))
            {
                ws.SetStyleInner(rowNum, 0, styleCashe[s]);
            }
            else
            {
                ExcelXfs st = this.CellXfs[s];
                int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                styleCashe.Add(s, newId);
                ws.SetStyleInner(rowNum, 0, newId);
            }
        }

        //Update individual cells 
        CellStoreEnumerator<ExcelValue>? cse2 =
            new CellStoreEnumerator<ExcelValue>(ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);

        while (cse2.Next())
        {
            int s = cse2.Value._styleId;

            if (s == 0)
            {
                continue;
            }

            if (styleCashe.ContainsKey(s))
            {
                ws.SetStyleInner(cse2.Row, cse2.Column, styleCashe[s]);
            }
            else
            {
                ExcelXfs st = this.CellXfs[s];
                int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                styleCashe.Add(s, newId);
                ws.SetStyleInner(cse2.Row, cse2.Column, newId);
            }
        }

        //Update cells with styled rows
        cse2 = new CellStoreEnumerator<ExcelValue>(ws._values, 0, 1, 0, address._toCol);

        while (cse2.Next())
        {
            if (cse2.Value._styleId == 0)
            {
                continue;
            }

            for (int r = address._fromRow; r <= address._toRow; r++)
            {
                if (!ws.ExistsStyleInner(r, cse2.Column))
                {
                    int s = cse2.Value._styleId;

                    if (styleCashe.ContainsKey(s))
                    {
                        ws.SetStyleInner(r, cse2.Column, styleCashe[s]);
                    }
                    else
                    {
                        ExcelXfs st = this.CellXfs[s];
                        int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                        styleCashe.Add(s, newId);
                        ws.SetStyleInner(r, cse2.Column, newId);
                    }
                }
            }
        }
    }

    private void SetStyleFullColumn(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, Dictionary<int, int> styleCashe)
    {
        ExcelColumn column;

        int col = address.Start.Column,
            row = 0;

        bool isNew;

        //Get the startcolumn
        object o = null;

        if (!ws.ExistsValueInner(0, address.Start.Column, ref o))
        {
            column = ws.Column(address.Start.Column);
            isNew = true;
        }
        else
        {
            column = (ExcelColumn)o;
            isNew = false;
        }

        int prevColumMax = column.ColumnMax;

        while (column.ColumnMin <= address.End.Column)
        {
            if (column.ColumnMin > prevColumMax + 1)
            {
                ExcelColumn? newColumn = ws.Column(prevColumMax + 1);
                newColumn.ColumnMax = column.ColumnMin - 1;
                this.AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
            }

            if (column.ColumnMax > address.End.Column)
            {
                ExcelColumn? newCol = ws.CopyColumn(column, address.End.Column + 1, column.ColumnMax);
                column.ColumnMax = address.End.Column;
            }

            int s = ws.GetStyleInner(0, column.ColumnMin);
            this.AddNewStyleColumn(sender, e, ws, styleCashe, column, s);

            //index++;
            prevColumMax = column.ColumnMax;

            if (!ws._values.NextCell(ref row, ref col) || row > 0)
            {
                if (column._columnMax == address.End.Column)
                {
                    break;
                }

                if (isNew)
                {
                    column._columnMax = address.End.Column;
                }
                else
                {
                    ExcelColumn? newColumn = ws.Column(column._columnMax + 1);
                    newColumn.ColumnMax = address.End.Column;
                    this.AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
                    column = newColumn;
                }

                break;
            }
            else
            {
                column = ws.GetValueInner(0, col) as ExcelColumn;
            }
        }

        if (column._columnMax < address.End.Column)
        {
            ExcelColumn? newCol = ws.Column(column._columnMax + 1) as ExcelColumn;
            newCol._columnMax = address.End.Column;

            int s = ws.GetStyleInner(0, column.ColumnMin);

            if (styleCashe.ContainsKey(s))
            {
                ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
            }
            else
            {
                ExcelXfs st = this.CellXfs[s];
                int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                styleCashe.Add(s, newId);
                ws.SetStyleInner(0, column.ColumnMin, newId);
            }

            column._columnMax = address.End.Column;
        }

        //Set for individual cells in the span. We loop all cells here since the cells are sorted with columns first.
        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(ws._values, 1, address._fromCol, address._toRow, address._toCol);

        while (cse.Next())
        {
            if (cse.Column >= address.Start.Column && cse.Column <= address.End.Column && cse.Value._styleId != 0)
            {
                if (styleCashe.ContainsKey(cse.Value._styleId))
                {
                    ws.SetStyleInner(cse.Row, cse.Column, styleCashe[cse.Value._styleId]);
                }
                else
                {
                    ExcelXfs st = this.CellXfs[cse.Value._styleId];
                    int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                    styleCashe.Add(cse.Value._styleId, newId);
                    ws.SetStyleInner(cse.Row, cse.Column, newId);
                }
            }
        }

        if (!(address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns))
        {
            //Update cells with styled columns
            cse = new CellStoreEnumerator<ExcelValue>(ws._values, 1, 0, address._toRow, 0);

            while (cse.Next())
            {
                if (cse.Value._styleId == 0)
                {
                    continue;
                }

                for (int c = address._fromCol; c <= address._toCol; c++)
                {
                    if (!ws.ExistsStyleInner(cse.Row, c))
                    {
                        if (styleCashe.ContainsKey(cse.Value._styleId))
                        {
                            ws.SetStyleInner(cse.Row, c, styleCashe[cse.Value._styleId]);
                        }
                        else
                        {
                            ExcelXfs st = this.CellXfs[cse.Value._styleId];
                            int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                            styleCashe.Add(cse.Value._styleId, newId);
                            ws.SetStyleInner(cse.Row, c, newId);
                        }
                    }
                }
            }
        }
    }

    private void AddNewStyleColumn(StyleBase sender, StyleChangeEventArgs e, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, int s)
    {
        if (styleCashe.ContainsKey(s))
        {
            ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
        }
        else
        {
            ExcelXfs st = this.CellXfs[s];
            int newId = st.GetNewID(this.CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
            styleCashe.Add(s, newId);
            ws.SetStyleInner(0, column.ColumnMin, newId);
        }
    }

    internal static int GetStyleId(ExcelWorksheet ws, int row, int col)
    {
        int v = 0;

        if (ws.ExistsStyleInner(row, col, ref v))
        {
            return v;
        }
        else
        {
            if (ws.ExistsStyleInner(row, 0, ref v)) //First Row
            {
                return v;
            }
            else // then column
            {
                if (ws.ExistsStyleInner(0, col, ref v))
                {
                    return v;
                }
                else
                {
                    int r = 0,
                        c = col;

                    if (ws._values.PrevCell(ref r, ref c))
                    {
                        //var column=ws.GetValueInner(0,c) as ExcelColumn;
                        ExcelValue val = ws._values.GetValue(0, c);
                        ExcelColumn? column = (ExcelColumn)val._value;

                        if (column != null && column.ColumnMax >= col) //Fixes issue 15174
                        {
                            //return ws.GetStyleInner(0, c);
                            return val._styleId;
                        }
                        else
                        {
                            return 0;
                        }
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Handles property changes on Named styles.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    /// <returns></returns>
    internal int NamedStylePropertyChange(StyleBase sender, StyleChangeEventArgs e)
    {
        int index = this.NamedStyles.FindIndexById(e.Address);

        if (index >= 0)
        {
            if (e.StyleClass == eStyleClass.Font
                && (e.StyleProperty == eStyleProperty.Name || e.StyleProperty == eStyleProperty.Size)
                && this.NamedStyles[index].BuildInId == 0)
            {
                foreach (ExcelWorksheet? ws in this._wb.Worksheets)
                {
                    ws.NormalStyleChange();
                }
            }

            int newId = this.CellStyleXfs[this.NamedStyles[index].StyleXfId].GetNewID(this.CellStyleXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
            int prevIx = this.NamedStyles[index].StyleXfId;
            this.NamedStyles[index].StyleXfId = newId;
            this.NamedStyles[index].Style.Index = newId;

            this.NamedStyles[index].XfId = int.MinValue;

            foreach (ExcelXfs? style in this.CellXfs)
            {
                if (style.XfId == prevIx)
                {
                    style.XfId = newId;
                }
            }
        }

        return 0;
    }

    /// <summary>
    /// Contains all numberformats for the package
    /// </summary>
    public ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats = new ExcelStyleCollection<ExcelNumberFormatXml>();

    /// <summary>
    /// Contains all font styles for the package
    /// </summary>
    public ExcelStyleCollection<ExcelFontXml> Fonts = new ExcelStyleCollection<ExcelFontXml>();

    /// <summary>
    /// Contains all fill styles for the package
    /// </summary>
    public ExcelStyleCollection<ExcelFillXml> Fills = new ExcelStyleCollection<ExcelFillXml>();

    /// <summary>
    /// Contain all border styles for the package
    /// </summary>
    public ExcelStyleCollection<ExcelBorderXml> Borders = new ExcelStyleCollection<ExcelBorderXml>();

    /// <summary>
    /// Contain all named cell styles for the package
    /// </summary>
    public ExcelStyleCollection<ExcelXfs> CellStyleXfs = new ExcelStyleCollection<ExcelXfs>();

    /// <summary>
    /// Contain all cell styles for the package
    /// </summary>
    public ExcelStyleCollection<ExcelXfs> CellXfs = new ExcelStyleCollection<ExcelXfs>();

    /// <summary>
    /// Contain all named styles for the package
    /// </summary>
    public ExcelStyleCollection<ExcelNamedStyleXml> NamedStyles = new ExcelStyleCollection<ExcelNamedStyleXml>();

    /// <summary>
    /// Contain all table styles for the package. Tables styles can be used to customly format tables and pivot tables.
    /// </summary>
    public ExcelNamedStyleCollection<ExcelTableNamedStyleBase> TableStyles = new ExcelNamedStyleCollection<ExcelTableNamedStyleBase>();

    /// <summary>
    /// Contain all slicer styles for the package. Tables styles can be used to customly format tables and pivot tables.
    /// </summary>
    public ExcelNamedStyleCollection<ExcelSlicerNamedStyle> SlicerStyles = new ExcelNamedStyleCollection<ExcelSlicerNamedStyle>();

    /// <summary>
    /// Contain differential formatting styles for the package. This collection does not contain style records for slicers.
    /// </summary>
    public ExcelStyleCollection<ExcelDxfStyleBase> Dxfs = new ExcelStyleCollection<ExcelDxfStyleBase>();

    internal ExcelStyleCollection<ExcelDxfStyleBase> DxfsSlicers = new ExcelStyleCollection<ExcelDxfStyleBase>();
    internal Dictionary<string, XmlNode> _slicerTableStyleNodes = new Dictionary<string, XmlNode>(StringComparer.InvariantCultureIgnoreCase);

    internal static string Id
    {
        get { return ""; }
    }

    /// <summary>
    /// Creates a named style that can be applied to cells in the worksheet.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <returns>A named style object that can be custumized</returns>
    public ExcelNamedStyleXml CreateNamedStyle(string name)
    {
        return this.CreateNamedStyle(name, null);
    }

    /// <summary>
    /// Creates a named style that can be applied to cells in the worksheet.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="Template">A template style</param>
    /// <returns>A named style object that can be custumized</returns>
    public ExcelNamedStyleXml CreateNamedStyle(string name, ExcelStyle Template)
    {
        if (this._wb.Styles.NamedStyles.ExistsKey(name))
        {
            throw new Exception(string.Format("Key {0} already exists in collection", name));
        }

        ExcelNamedStyleXml style = new(this.NameSpaceManager, this);

        int xfIdCopy,
            positionID;

        ExcelStyles styles;
        bool isTemplateNamedStyle;

        if (Template == null)
        {
            ExcelNamedStyleXml? ns = this._wb.Styles.GetNormalStyle();

            if (ns != null)
            {
                xfIdCopy = ns.StyleXfId;
            }
            else
            {
                xfIdCopy = -1;
            }

            positionID = -1;
            styles = this;
            isTemplateNamedStyle = true;
        }
        else
        {
            isTemplateNamedStyle = Template.PositionID == -1;

            if (Template.PositionID < 0 && Template.Styles == this)
            {
                xfIdCopy = Template.Index;

                positionID = Template.PositionID;
                styles = this;
            }
            else
            {
                xfIdCopy = Template.Index;
                positionID = -1;
                styles = Template.Styles;
            }
        }

        //Clone namedstyle
        if (xfIdCopy >= 0)
        {
            int styleXfId = this.CloneStyle(styles, xfIdCopy, true, false, isTemplateNamedStyle);
            this.CellStyleXfs[styleXfId].XfId = this.CellStyleXfs.Count - 1;
            style.Style = new ExcelStyle(this, this.NamedStylePropertyChange, positionID, name, styleXfId);
            style.StyleXfId = styleXfId;
        }
        else
        {
            style.Style = new ExcelStyle(this, this.NamedStylePropertyChange, positionID, name, 0);
            style.StyleXfId = 0;
        }

        style.Name = name;
        int ix = this._wb.Styles.NamedStyles.Add(style.Name, style);
        style.Style.SetIndex(ix);

        return style;
    }

    /// <summary>
    /// Creates a tables style only visible for pivot tables and with elements specific to pivot tables.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <returns>The table style object</returns>
    public ExcelPivotTableNamedStyle CreatePivotTableStyle(string name)
    {
        this.ValidateTableStyleName(name);
        XmlElement? node = (XmlElement)this.CreateNode("d:tableStyles/d:tableStyle", false, true);
        node.SetAttribute("table", "0");
        ExcelPivotTableNamedStyle? s = new ExcelPivotTableNamedStyle(this.NameSpaceManager, node, this) { Name = name };

        this.TableStyles.Add(name, s);

        return s;
    }

    /// <summary>
    /// Creates a tables style only visible for pivot tables and with elements specific to pivot tables.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The built-in table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelPivotTableNamedStyle CreatePivotTableStyle(string name, PivotTableStyles templateStyle)
    {
        ExcelPivotTableNamedStyle? s = this.CreatePivotTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a tables style only visible for pivot tables and with elements specific to pivot tables.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelPivotTableNamedStyle CreatePivotTableStyle(string name, ExcelTableNamedStyleBase templateStyle)
    {
        ExcelPivotTableNamedStyle? s = this.CreatePivotTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a tables style only visible for tables and with elements specific to pivot tables.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <returns>The table style object</returns>
    public ExcelTableNamedStyle CreateTableStyle(string name)
    {
        this.ValidateTableStyleName(name);
        XmlElement? node = (XmlElement)this.CreateNode("d:tableStyles/d:tableStyle", false, true);
        node.SetAttribute("pivot", "0");
        ExcelTableNamedStyle? s = new ExcelTableNamedStyle(this.NameSpaceManager, node, this) { Name = name };

        this.TableStyles.Add(name, s);

        return s;
    }

    /// <summary>
    /// Creates a tables style only visible for tables and with elements specific to pivot tables.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The built-in table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelTableNamedStyle CreateTableStyle(string name, TableStyles templateStyle)
    {
        ExcelTableNamedStyle? s = this.CreateTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a tables style only visible for tables and with elements specific to pivot tables.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelTableNamedStyle CreateTableStyle(string name, ExcelTableNamedStyleBase templateStyle)
    {
        ExcelTableNamedStyle? s = this.CreateTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a tables visible for tables and pivot tables and with elements for both.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <returns>The table style object</returns>
    public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name)
    {
        this.ValidateTableStyleName(name);
        XmlElement? node = (XmlElement)this.CreateNode("d:tableStyles/d:tableStyle", false, true);
        ExcelTableAndPivotTableNamedStyle? s = new ExcelTableAndPivotTableNamedStyle(this.NameSpaceManager, node, this) { Name = name };

        this.TableStyles.Add(name, s);

        return s;
    }

    /// <summary>
    /// Creates a tables visible for tables and pivot tables and with elements for both.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The built-in table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name, TableStyles templateStyle)
    {
        if (templateStyle == Table.TableStyles.Custom)
        {
            throw new ArgumentException("Cant use template style Custom. To use a custom style, please use the ´PivotTableStyles´ overload of this method.",
                                        nameof(templateStyle));
        }

        ExcelTableAndPivotTableNamedStyle? s = this.CreateTableAndPivotTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a tables visible for tables and pivot tables and with elements for both.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The built-in pivot table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name, PivotTableStyles templateStyle)
    {
        if (templateStyle == PivotTableStyles.Custom)
        {
            throw new
                ArgumentException("Cant use template style Custom. To use a custom style, please use the ´ExcelTableNamedStyleBase´ overload of this method.",
                                  nameof(templateStyle));
        }

        ExcelTableAndPivotTableNamedStyle? s = this.CreateTableAndPivotTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a tables visible for tables and pivot tables and with elements for both.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The table style to use as a template for this custom style</param>
    /// <returns>The table style object</returns>
    public ExcelTableAndPivotTableNamedStyle CreateTableAndPivotTableStyle(string name, ExcelTableNamedStyleBase templateStyle)
    {
        ExcelTableAndPivotTableNamedStyle? s = this.CreateTableAndPivotTableStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Creates a custom slicer style.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <returns>The slicer style object</returns>
    public ExcelSlicerNamedStyle CreateSlicerStyle(string name)
    {
        this.ValidateTableStyleName(name);

        //Create the matching table style
        XmlElement? tableStyleNode = (XmlElement)this.CreateNode("d:tableStyles/d:tableStyle", false, true);
        tableStyleNode.SetAttribute("table", "0");
        tableStyleNode.SetAttribute("pivot", "0");
        tableStyleNode.SetAttribute("name", name);
        this._slicerTableStyleNodes.Add(name, tableStyleNode);

        //The dxfs collection must be created before the slicer styles collection
        this.GetOrCreateExtLstSubNode(ExtLstUris.SlicerStylesDxfCollectionUri, "x14");

        XmlNode? extNode = this.GetOrCreateExtLstSubNode(ExtLstUris.SlicerStylesUri, "x14");
        XmlHelper? extHelper = XmlHelperFactory.Create(this.NameSpaceManager, extNode);

        if (extNode.ChildNodes.Count == 0)
        {
            XmlElement? slicersNode = (XmlElement)extHelper.CreateNode("x14:slicerStyles", false, true);
            slicersNode.SetAttribute("defaultSlicerStyle", "SlicerStyleLight1"); //defaultSlicerStyle is required
            extHelper.TopNode = slicersNode;
        }
        else
        {
            extHelper.TopNode = extNode.FirstChild;
        }

        XmlElement? node = (XmlElement)extHelper.CreateNode("x14:slicerStyle", false, true);

        ExcelSlicerNamedStyle? s = new ExcelSlicerNamedStyle(this.NameSpaceManager, node, tableStyleNode, this) { Name = name };

        this.SlicerStyles.Add(name, s);

        return s;
    }

    /// <summary>
    /// Creates a custom slicer style.
    /// </summary>
    /// <param name="name">The name of the style</param>
    /// <param name="templateStyle">The slicer style to use as a template for this custom style</param>
    /// <returns>The slicer style object</returns>
    public ExcelSlicerNamedStyle CreateSlicerStyle(string name, eSlicerStyle templateStyle)
    {
        if (templateStyle == eSlicerStyle.Custom)
        {
            throw new
                ArgumentException("Cant use template style Custom. To use a custom style, please use the ´ExcelSlicerNamedStyle´ overload of this method.",
                                  nameof(templateStyle));
        }

        ExcelSlicerNamedStyle? s = this.CreateSlicerStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    HashSet<string> tableStyleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    private void ValidateTableStyleName(string name)
    {
        if (this.tableStyleNames.Count == 0)
        {
            Enum.GetNames(typeof(TableStyles)).Select(x => this.tableStyleNames.Add("TableStyle" + x));
            Enum.GetNames(typeof(PivotTableStyles)).Select(x => this.tableStyleNames.Add("PivotTableStyle" + x));
            Enum.GetNames(typeof(eSlicerStyle)).Select(x => this.tableStyleNames.Add("SlicerStyle" + x));
        }

        if (this.tableStyleNames.Contains(name) || this.TableStyles.ExistsKey(name) || this.SlicerStyles.ExistsKey(name))
        {
            throw new ArgumentException($"Table style name is not unique : {name}", "name");
        }
    }

    /// <summary>
    /// Creates a custom named slicer style from another style.
    /// </summary>
    /// <param name="name">The name of the style.</param>
    /// <param name="templateStyle">The slicer style to us as template.</param>
    /// <returns></returns>
    public ExcelSlicerNamedStyle CreateSlicerStyle(string name, ExcelSlicerNamedStyle templateStyle)
    {
        ExcelSlicerNamedStyle? s = this.CreateSlicerStyle(name);
        s.SetFromTemplate(templateStyle);

        return s;
    }

    /// <summary>
    /// Update the changes to the Style.Xml file inside the package.
    /// This will remove any unused styles from the collections.
    /// </summary>
    public void UpdateXml()
    {
        this.RemoveUnusedStyles();

        int normalIx = this.GetNormalStyleIndex();

        this.UpdateNumberFormatXml(normalIx);
        this.UpdateFontXml(normalIx);
        this.UpdateFillXml();
        this.UpdateBorderXml();
        this.UpdateNamedStylesAndXfs(normalIx);

        DxfStyleHandler.UpdateDxfXml(this._wb);
    }

    private void UpdateNamedStylesAndXfs(int normalIx)
    {
        //Create the cellStyleXfs element            
        XmlNode styleXfsNode = this.GetNode(CellStyleXfsPath);

        if (styleXfsNode == null)
        {
            if (this.CellStyleXfs.Count > 0)
            {
                styleXfsNode = this.CreateNode(CellStyleXfsPath);
            }
        }
        else
        {
            styleXfsNode?.RemoveAll();
        }

        //NamedStyles
        int count = normalIx > -1 ? 1 : 0; //If we have a normal style, we make sure it's added first.

        XmlNode cellStyleNode = this.GetNode(CellStylesPath);

        if (cellStyleNode == null)
        {
            if (this.NamedStyles.Count > 0)
            {
                cellStyleNode = this.CreateNode(CellStylesPath);
            }
        }
        else
        {
            cellStyleNode.RemoveAll();
        }

        XmlNode cellXfsNode = this.GetNode(CellXfsPath);
        cellXfsNode.RemoveAll();
        int xfsCount = 0;

        if (this.CellStyleXfs.Count > 0)
        {
            if (normalIx >= 0)
            {
                this.NamedStyles[normalIx].newID = 0;
                this.AddNamedStyle(0, styleXfsNode, cellXfsNode, this.NamedStyles[normalIx]);

                cellXfsNode.AppendChild(this.CellStyleXfs[this.NamedStyles[normalIx].StyleXfId]
                                            .CreateXmlNode(this._styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
            }
            else
            {
                styleXfsNode.AppendChild(this.CellStyleXfs[0].CreateXmlNode(this._styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
                this.CellStyleXfs[0].newID = 0;
                cellXfsNode.AppendChild(this.CellStyleXfs[0].CreateXmlNode(this._styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
            }

            xfsCount++;
        }

        foreach (ExcelNamedStyleXml style in this.NamedStyles)
        {
            if (style.BuildInId != 0)
            {
                this.AddNamedStyle(count++, styleXfsNode, cellXfsNode, style);
            }
            else
            {
                style.newID = 0;
            }

            cellStyleNode.AppendChild(style.CreateXmlNode(this._styleXml.CreateElement("cellStyle", ExcelPackage.schemaMain)));
        }

        if (cellStyleNode != null)
        {
            XmlElement? cellStyleElement = cellStyleNode as XmlElement;
            cellStyleElement.SetAttribute("count", cellStyleElement.ChildNodes.Count.ToString(CultureInfo.InvariantCulture));
        }

        if (styleXfsNode != null)
        {
            XmlElement? styleXfsElement = styleXfsNode as XmlElement;
            styleXfsElement.SetAttribute("count", styleXfsElement.ChildNodes.Count.ToString(CultureInfo.InvariantCulture));
        }

        //CellStyle
        int xfix = 0;

        foreach (ExcelXfs xf in this.CellXfs)
        {
            if (xf.useCnt > 0 && !(xfix == 0 && normalIx >= 0))
            {
                cellXfsNode.AppendChild(xf.CreateXmlNode(this._styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                xf.newID = xfsCount;
                xfsCount++;
            }

            xfix++;
        }

        (cellXfsNode as XmlElement).SetAttribute("count", xfsCount.ToString(CultureInfo.InvariantCulture));
    }

    private void UpdateBorderXml()
    {
        //Borders
        int count = 0;
        XmlNode bordersNode = this.GetNode(BordersPath);
        bordersNode.RemoveAll();
        this.Borders[0].useCnt = 1; //Must exist blank;

        foreach (ExcelBorderXml border in this.Borders)
        {
            if (border.useCnt > 0)
            {
                bordersNode.AppendChild(border.CreateXmlNode(this._styleXml.CreateElement("border", ExcelPackage.schemaMain)));
                border.newID = count;
                count++;
            }
        }

        (bordersNode as XmlElement).SetAttribute("count", count.ToString());
    }

    private int UpdateFillXml()
    {
        //Fills
        int count = 0;
        XmlNode fillsNode = this.GetNode(FillsPath);
        fillsNode.RemoveAll();
        this.Fills[0].useCnt = 1; //Must exist (none);  
        this.Fills[1].useCnt = 1; //Must exist (gray125);

        foreach (ExcelFillXml fill in this.Fills)
        {
            if (fill.useCnt > 0)
            {
                fillsNode.AppendChild(fill.CreateXmlNode(this._styleXml.CreateElement("fill", ExcelPackage.schemaMain)));
                fill.newID = count;
                count++;
            }
        }

        (fillsNode as XmlElement).SetAttribute("count", count.ToString());

        return count;
    }

    internal int GetNormalStyleIndex()
    {
        int normalIx = this.NamedStyles.FindIndexByBuildInId(0);

        if (normalIx < 0)
        {
            normalIx = this.NamedStyles.FindIndexById("normal");
        }

        return normalIx;
    }

    private void UpdateFontXml(int normalIx)
    {
        //Font
        int count = 0;
        XmlNode fntNode = this.GetNode(FontsPath);
        fntNode.RemoveAll();
        int nfIx = -1;

        //Normal should be first in the collection
        if (this.NamedStyles.Count > 0 && normalIx >= 0 && this.NamedStyles[normalIx].Style.Font.Index >= 0)
        {
            nfIx = this.NamedStyles[normalIx].Style.Font.Index;
            ExcelFontXml fnt = this.Fonts[nfIx];
            fntNode.AppendChild(fnt.CreateXmlNode(this._styleXml.CreateElement("font", ExcelPackage.schemaMain)));
            fnt.newID = count++;
        }

        int ix = 0;

        foreach (ExcelFontXml fnt in this.Fonts)
        {
            if (fnt.useCnt > 0 && ix != nfIx)
            {
                fntNode.AppendChild(fnt.CreateXmlNode(this._styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                fnt.newID = count;
                count++;
            }

            ix++;
        }

        (fntNode as XmlElement).SetAttribute("count", count.ToString());
    }

    private void UpdateNumberFormatXml(int normalIx)
    {
        //NumberFormat
        XmlNode nfNode = this.GetNode(NumberFormatsPath);

        if (nfNode == null)
        {
            nfNode = this.CreateNode(NumberFormatsPath, true);
        }
        else
        {
            nfNode.RemoveAll();
        }

        int count = 0;

        if (this.NamedStyles.Count > 0 && normalIx >= 0 && this.NamedStyles[normalIx].Style.Numberformat.NumFmtID >= 164)
        {
            ExcelNumberFormatXml nf = this.NumberFormats[this.NumberFormats.FindIndexById(this.NamedStyles[normalIx].Style.Numberformat.Id)];
            nfNode.AppendChild(nf.CreateXmlNode(this._styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
            nf.newID = count++;
        }

        //Add pivot Table formatting.
        foreach (ExcelWorksheet? ws in this._wb.Worksheets)
        {
            if (!(ws is ExcelChartsheet) && ws.HasLoadedPivotTables)
            {
                foreach (ExcelPivotTable? pt in ws.PivotTables)
                {
                    foreach (ExcelPivotTableField? f in pt.Fields)
                    {
                        f.NumFmtId = this.GetNumFormatId(f.Format);
                        f.Cache.NumFmtId = this.GetNumFormatId(f.Cache.Format);
                    }

                    foreach (ExcelPivotTableDataField? df in pt.DataFields)
                    {
                        if (df.NumFmtId < 0 && df.Field.NumFmtId.HasValue)
                        {
                            df.NumFmtId = df.Field.NumFmtId.Value;
                        }
                    }
                }
            }
        }

        foreach (ExcelNumberFormatXml nf in this.NumberFormats)
        {
            if (!nf.BuildIn /*&& nf.newID<0*/) //Buildin formats are not updated.
            {
                nfNode.AppendChild(nf.CreateXmlNode(this._styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                nf.newID = count;
                count++;
            }
        }

        (nfNode as XmlElement).SetAttribute("count", count.ToString());
    }

    private int? GetNumFormatId(string format)
    {
        if (string.IsNullOrEmpty(format))
        {
            return null;
        }
        else
        {
            ExcelNumberFormatXml nf = null;

            if (this.NumberFormats.FindById(format, ref nf))
            {
                return nf.NumFmtId;
            }
            else
            {
                int id = this.NumberFormats.NextId++;
                ExcelNumberFormatXml? item = new ExcelNumberFormatXml(this.NameSpaceManager, false) { Format = format, NumFmtId = id };

                this.NumberFormats.Add(format, item);

                return id;
            }
        }
    }

    private void AddNamedStyle(int id, XmlNode styleXfsNode, XmlNode cellXfsNode, ExcelNamedStyleXml style)
    {
        ExcelXfs? styleXfs = this.CellStyleXfs[style.StyleXfId];
        styleXfsNode.AppendChild(styleXfs.CreateXmlNode(this._styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
        styleXfs.newID = id;
        styleXfs.XfId = style.StyleXfId;

        if (style.XfId >= 0)
        {
            style.XfId = this.CellXfs[style.XfId].newID;
        }
        else
        {
            style.XfId = 0;
        }
    }

    private void RemoveUnusedStyles()
    {
        this.CellXfs[0].useCnt = 1; //First item is allways used.

        foreach (ExcelWorksheet sheet in this._wb.Worksheets)
        {
            if (sheet is ExcelChartsheet)
            {
                continue;
            }

            CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(sheet._values);

            while (cse.Next())
            {
                int v = cse.Value._styleId;

                if (v >= 0)
                {
                    this.CellXfs[v].useCnt++;
                }
            }
        }

        foreach (ExcelNamedStyleXml ns in this.NamedStyles)
        {
            this.CellStyleXfs[ns.StyleXfId].useCnt++;
        }

        foreach (ExcelXfs xf in this.CellXfs)
        {
            if (xf.useCnt > 0)
            {
                if (xf.FontId >= 0)
                {
                    this.Fonts[xf.FontId].useCnt++;
                }

                if (xf.FillId >= 0)
                {
                    this.Fills[xf.FillId].useCnt++;
                }

                if (xf.BorderId >= 0)
                {
                    this.Borders[xf.BorderId].useCnt++;
                }
            }
        }

        foreach (ExcelXfs xf in this.CellStyleXfs)
        {
            if (xf.useCnt > 0)
            {
                if (xf.FontId >= 0)
                {
                    this.Fonts[xf.FontId].useCnt++;
                }

                if (xf.FillId >= 0)
                {
                    this.Fills[xf.FillId].useCnt++;
                }

                if (xf.BorderId >= 0)
                {
                    this.Borders[xf.BorderId].useCnt++;
                }
            }
        }
    }

    internal int GetStyleIdFromName(string Name)
    {
        int i = this.NamedStyles.FindIndexById(Name);

        if (i >= 0)
        {
            int id = this.NamedStyles[i].XfId;

            if (id < 0)
            {
                int styleXfId = this.NamedStyles[i].StyleXfId;
                ExcelXfs newStyle = this.CellStyleXfs[styleXfId].Copy();
                newStyle.XfId = styleXfId;
                id = this.CellXfs.FindIndexById(newStyle.Id);

                if (id < 0)
                {
                    id = this.CellXfs.Add(newStyle.Id, newStyle);
                }

                this.NamedStyles[i].XfId = id;
            }

            return id;
        }
        else
        {
            return 0;

            //throw(new Exception("Named style does not exist"));                    
        }
    }

    #region XmlHelpFunctions

    private static int GetXmlNodeInt(XmlNode node)
    {
        if (int.TryParse(GetXmlNode(node), out int i))
        {
            return i;
        }
        else
        {
            return 0;
        }
    }

    private static string GetXmlNode(XmlNode node)
    {
        if (node == null)
        {
            return "";
        }

        if (node.Value != null)
        {
            return node.Value;
        }
        else
        {
            return "";
        }
    }

    #endregion

    internal int CloneStyle(ExcelStyles style, int styleID)
    {
        return this.CloneStyle(style, styleID, false, false, false);
    }

    internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle)
    {
        return this.CloneStyle(style, styleID, isNamedStyle, false, isNamedStyle);
    }

    internal int CloneStyle(ExcelStyles style, int styleID, bool isNamedStyle, bool allwaysAddCellXfs, bool isCellStyleXfs)
    {
        lock (style)
        {
            ExcelXfs xfs;

            if (isCellStyleXfs)
            {
                xfs = style.CellStyleXfs[styleID];
            }
            else
            {
                xfs = style.CellXfs[styleID];
            }

            ExcelXfs newXfs = xfs.Copy(this);

            //Numberformat
            if (xfs.NumberFormatId > 0)
            {
                //rake36: Two problems here...
                //rake36:  1. the first time through when format stays equal to String.Empty, it adds a string.empty to the list of Number Formats
                //rake36:  2. when adding a second sheet, if the numberformatid == 164, it finds the 164 added by previous sheets but was using the array index
                //rake36:      for the numberformatid

                string format = string.Empty;

                foreach (ExcelNumberFormatXml? fmt in style.NumberFormats)
                {
                    if (fmt.NumFmtId == xfs.NumberFormatId)
                    {
                        format = fmt.Format;

                        break;
                    }
                }

                //rake36: Don't add another format if it's blank
                if (!String.IsNullOrEmpty(format))
                {
                    int ix = this.NumberFormats.FindIndexById(format);

                    if (ix < 0)
                    {
                        ExcelNumberFormatXml? item =
                            new ExcelNumberFormatXml(this.NameSpaceManager) { Format = format, NumFmtId = this.NumberFormats.NextId++ };

                        this.NumberFormats.Add(format, item);

                        //rake36: Use the just added format id
                        newXfs.NumberFormatId = item.NumFmtId;
                    }
                    else
                    {
                        //rake36: Use the format id defined by the index... not the index itself
                        newXfs.NumberFormatId = this.NumberFormats[ix].NumFmtId;
                    }
                }
            }

            //Font
            if (xfs.FontId > -1)
            {
                int ix = this.Fonts.FindIndexById(xfs.Font.Id);

                if (ix < 0)
                {
                    ExcelFontXml item = style.Fonts[xfs.FontId].Copy();
                    ix = this.Fonts.Add(xfs.Font.Id, item);
                }

                newXfs.FontId = ix;
            }

            //Border
            if (xfs.BorderId > -1)
            {
                int ix = this.Borders.FindIndexById(xfs.Border.Id);

                if (ix < 0)
                {
                    ExcelBorderXml item = style.Borders[xfs.BorderId].Copy();
                    ix = this.Borders.Add(xfs.Border.Id, item);
                }

                newXfs.BorderId = ix;
            }

            //Fill
            if (xfs.FillId > -1)
            {
                int ix = this.Fills.FindIndexById(xfs.Fill.Id);

                if (ix < 0)
                {
                    ExcelFillXml? item = style.Fills[xfs.FillId].Copy();
                    ix = this.Fills.Add(xfs.Fill.Id, item);
                }

                newXfs.FillId = ix;
            }

            //Named style reference
            if (xfs.XfId > 0)
            {
                if (style._wb != this._wb && allwaysAddCellXfs == false) //Not the same workbook, copy the namedstyle to the workbook or match the id
                {
                    Dictionary<int, ExcelNamedStyleXml>? nsFind = style.NamedStyles.ToDictionary(d => d.StyleXfId);

                    if (nsFind.ContainsKey(xfs.XfId))
                    {
                        ExcelNamedStyleXml? st = nsFind[xfs.XfId];

                        if (this.NamedStyles.ExistsKey(st.Name))
                        {
                            newXfs.XfId = this.NamedStyles.FindIndexById(st.Name);
                        }
                        else
                        {
                            ExcelNamedStyleXml? ns = this.CreateNamedStyle(st.Name, st.Style);
                            newXfs.XfId = this.NamedStyles.Count - 1;
                        }
                    }
                }
                else
                {
                    string? id = style.CellStyleXfs[xfs.XfId].Id;
                    int newId = this.CellStyleXfs.FindIndexById(id);

                    if (newId >= 0)
                    {
                        newXfs.XfId = newId;
                    }
                }
            }

            int index;

            if (isNamedStyle && allwaysAddCellXfs == false)
            {
                index = this.CellStyleXfs.Add(newXfs.Id, newXfs);
            }
            else
            {
                if (allwaysAddCellXfs)
                {
                    index = this.CellXfs.Add(newXfs.Id, newXfs);
                }
                else
                {
                    index = this.CellXfs.FindIndexById(newXfs.Id);

                    if (index < 0)
                    {
                        index = this.CellXfs.Add(newXfs.Id, newXfs);
                    }
                }
            }

            return index;
        }
    }

    internal ExcelDxfStyleLimitedFont GetDxfLimitedFont(int? dxfId)
    {
        if (dxfId.HasValue && dxfId < this.Dxfs.Count)
        {
            return this.Dxfs[dxfId.Value].ToDxfLimitedStyle();
        }
        else
        {
            return new ExcelDxfStyleLimitedFont(this.NameSpaceManager, null, this, null);
        }
    }

    internal ExcelDxfStyle GetDxf(int? dxfId, Action<eStyleClass, eStyleProperty, object> callback)
    {
        if (dxfId.HasValue && dxfId < this.Dxfs.Count)
        {
            return this.Dxfs[dxfId.Value].ToDxfStyle();
        }
        else
        {
            return new ExcelDxfStyle(this.NameSpaceManager, null, this, callback);
        }
    }

    internal ExcelDxfSlicerStyle GetDxfSlicer(int? dxfId, Action<eStyleClass, eStyleProperty, object> callback)
    {
        if (dxfId.HasValue && dxfId < this.Dxfs.Count)
        {
            return this.Dxfs[dxfId.Value].ToDxfSlicerStyle();
        }
        else
        {
            return new ExcelDxfSlicerStyle(this.NameSpaceManager, null, this, callback);
        }
    }

    internal ExcelDxfTableStyle GetDxfTable(int? dxfId, Action<eStyleClass, eStyleProperty, object> callback)
    {
        if (dxfId.HasValue && dxfId < this.Dxfs.Count)
        {
            return this.Dxfs[dxfId.Value].ToDxfTableStyle();
        }
        else
        {
            return new ExcelDxfTableStyle(this.NameSpaceManager, null, this, callback);
        }
    }

    internal ExcelDxfSlicerStyle GetDxfSlicer(int? dxfId)
    {
        if (dxfId.HasValue && dxfId < this.DxfsSlicers.Count)
        {
            return this.DxfsSlicers[dxfId.Value].ToDxfSlicerStyle();
        }
        else
        {
            return new ExcelDxfSlicerStyle(this.NameSpaceManager, null, this, null);
        }
    }
}