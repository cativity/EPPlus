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
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml;

/// <summary>
/// Access to workbook view properties
/// </summary>
public class ExcelWorkbookView : XmlHelper
{
    private readonly ExcelWorkbook _wb;

    #region ExcelWorksheetView Constructor

    /// <summary>
    /// Creates a new ExcelWorkbookView which provides access to all the 
    /// view states of the worksheet.
    /// </summary>
    /// <param name="ns"></param>
    /// <param name="node"></param>
    /// <param name="wb"></param>
    internal ExcelWorkbookView(XmlNamespaceManager ns, XmlNode node, ExcelWorkbook wb)
        : base(ns, node)
    {
        this.SchemaNodeOrder = wb.SchemaNodeOrder;
        this._wb = wb;
    }

    #endregion

    const string LEFT_PATH = "d:bookViews/d:workbookView/@xWindow";

    /// <summary>
    /// Position of the upper left corner of the workbook window. In twips.
    /// </summary>
    public int Left
    {
        get => this.GetXmlNodeInt(LEFT_PATH);
        internal set => this.SetXmlNodeString(LEFT_PATH, value.ToString());
    }

    const string TOP_PATH = "d:bookViews/d:workbookView/@yWindow";

    /// <summary>
    /// Position of the upper left corner of the workbook window. In twips.
    /// </summary>
    public int Top
    {
        get => this.GetXmlNodeInt(TOP_PATH);
        internal set => this.SetXmlNodeString(TOP_PATH, value.ToString());
    }

    const string WIDTH_PATH = "d:bookViews/d:workbookView/@windowWidth";

    /// <summary>
    /// Width of the workbook window. In twips.
    /// </summary>
    public int Width
    {
        get => this.GetXmlNodeInt(WIDTH_PATH);
        internal set => this.SetXmlNodeString(WIDTH_PATH, value.ToString());
    }

    const string HEIGHT_PATH = "d:bookViews/d:workbookView/@windowHeight";

    /// <summary>
    /// Height of the workbook window. In twips.
    /// </summary>
    public int Height
    {
        get => this.GetXmlNodeInt(HEIGHT_PATH);
        internal set => this.SetXmlNodeString(HEIGHT_PATH, value.ToString());
    }

    const string MINIMIZED_PATH = "d:bookViews/d:workbookView/@minimized";

    /// <summary>
    /// If true the the workbook window is minimized.
    /// </summary>
    public bool Minimized
    {
        get => this.GetXmlNodeBool(MINIMIZED_PATH);
        set => this.SetXmlNodeString(MINIMIZED_PATH, value.ToString());
    }

    const string SHOWVERTICALSCROLL_PATH = "d:bookViews/d:workbookView/@showVerticalScroll";

    /// <summary>
    /// Show the vertical scrollbar
    /// </summary>
    public bool ShowVerticalScrollBar
    {
        get => this.GetXmlNodeBool(SHOWVERTICALSCROLL_PATH, true);
        set => this.SetXmlNodeBool(SHOWVERTICALSCROLL_PATH, value, true);
    }

    const string SHOWHORIZONTALSCR_PATH = "d:bookViews/d:workbookView/@showHorizontalScroll";

    /// <summary>
    /// Show the horizontal scrollbar
    /// </summary>
    public bool ShowHorizontalScrollBar
    {
        get => this.GetXmlNodeBool(SHOWHORIZONTALSCR_PATH, true);
        set => this.SetXmlNodeBool(SHOWHORIZONTALSCR_PATH, value, true);
    }

    const string SHOWSHEETTABS_PATH = "d:bookViews/d:workbookView/@showSheetTabs";

    /// <summary>
    /// Show or hide the sheet tabs
    /// </summary>
    public bool ShowSheetTabs
    {
        get => this.GetXmlNodeBool(SHOWSHEETTABS_PATH, true);
        set => this.SetXmlNodeBool(SHOWSHEETTABS_PATH, value, true);
    }

    /// <summary>
    /// Set the window position in twips
    /// </summary>
    /// <param name="left">Left coordinat</param>
    /// <param name="top">Top coordinat</param>
    /// <param name="width">Width in twips</param>
    /// <param name="height">Height in twips</param>
    public void SetWindowSize(int left, int top, int width, int height)
    {
        this.Left = left;
        this.Top = top;
        this.Width = width;
        this.Height = height;
    }

    const string ACTIVETAB_PATH = "d:bookViews/d:workbookView/@activeTab";

    /// <summary>
    /// The active worksheet in the workbook. Zero based.
    /// </summary>
    public int ActiveTab
    {
        get
        {
            int v = this.GetXmlNodeInt(ACTIVETAB_PATH);

            if (v < 0)
            {
                return this._wb._package._worksheetAdd;
            }
            else
            {
                return v;
            }
        }
        set
        {
            if (value < 0 || value >= this._wb.Worksheets.Count)
            {
                throw new InvalidOperationException("Value out of range");
            }

            this.SetXmlNodeString(ACTIVETAB_PATH, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string FirstSheet_PATH = "d:bookViews/d:workbookView/@firstSheet";

    /// <summary>
    /// The first visible worksheet in the worksheets collection. 
    /// </summary>
    internal int? FirstSheet
    {
        get => this.GetXmlNodeIntNull(FirstSheet_PATH);
        set => this.SetXmlNodeInt(FirstSheet_PATH, value);
    }
}