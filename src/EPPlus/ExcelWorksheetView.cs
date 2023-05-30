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

using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

/// <summary>
/// The state of the pane.
/// </summary>
public enum ePaneState
{
    /// <summary>
    /// Panes are frozen, but were not split being frozen.In this state, when the panes are unfrozen again, a single pane results, with no split. In this state, the split bars are not adjustable.
    /// </summary>
    Frozen,

    /// <summary>
    /// Frozen Split
    /// Panes are frozen and were split before being frozen. In this state, when the panes are unfrozen again, the split remains, but is adjustable.
    /// </summary>
    FrozenSplit,

    /// <summary>
    /// Panes are split, but not frozen.In this state, the split bars are adjustable by the user.
    /// </summary>
    Split
}

/// <summary>
/// The position of the pane.
/// </summary>
public enum ePanePosition
{
    /// <summary>
    /// Bottom Left Pane.
    /// Used when worksheet view has both vertical and horizontal splits.
    /// Also used when the worksheet is horizontaly split only, specifying this is the bottom pane.
    /// </summary>
    BottomLeft,

    /// <summary>
    /// Bottom Right Pane. 
    /// This property is only used when the worksheet has both vertical and horizontal splits.
    /// </summary>
    BottomRight,

    /// <summary>
    /// Top Left Pane.
    /// Used when worksheet view has both vertical and horizontal splits.
    /// Also used when the worksheet is horizontaly split only, specifying this is the top pane.
    /// </summary>
    TopLeft,

    /// <summary>
    /// Top Right Pane
    /// Used when the worksheet view has both vertical and horizontal splits.
    /// Also used when the worksheet is verticaly split only, specifying this is the right pane.
    /// </summary>
    TopRight
}

/// <summary>
/// Represents the different view states of the worksheet
/// </summary>
public class ExcelWorksheetView : XmlHelper
{
    /// <summary>
    /// Defines general properties for the panes, if the worksheet is frozen or split.
    /// </summary>
    public class ExcelWorksheetViewPaneSettings : XmlHelper
    {
        internal ExcelWorksheetViewPaneSettings(XmlNamespaceManager ns, XmlNode topNode)
            : base(ns, topNode)
        {
        }

        /// <summary>
        /// The state of the pane.
        /// </summary>
        public ePaneState State
        {
            get => this.GetXmlEnumNull<ePaneState>("@state", ePaneState.Split).Value;
            internal set => this.SetXmlNodeString("@state", value.ToEnumString());
        }

        /// <summary>
        /// The active pane
        /// </summary>
        public ePanePosition ActivePanePosition
        {
            get => this.GetXmlEnumNull<ePanePosition>("@activePane", ePanePosition.TopLeft).Value;
            set => this.SetXmlNodeString("@activePane", value.ToEnumString());
        }

        /// <summary>
        /// The horizontal position of the split. 1/20 of a point if the pane is split. Number of columns in the top pane if this pane is frozen.
        /// </summary>
        public double XSplit
        {
            get => this.GetXmlNodeDouble("@xSplit");
            set => this.SetXmlNodeDouble("@xSplit", value, false);
        }

        /// <summary>
        /// The vertical position of the split. 1/20 of a point if the pane is split. Number of rows in the left pane if this pane is frozen.
        /// </summary>
        public double YSplit
        {
            get => this.GetXmlNodeDouble("@ySplit");
            set => this.SetXmlNodeDouble("@ySplit", value, false);
        }

        /// <summary>
        /// 
        /// </summary>
        public string TopLeftCell
        {
            get => this.GetXmlNodeString("@topLeftCell");
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    this.DeleteNode("@topLeftCell");
                }
                else if (ExcelCellBase.IsValidCellAddress(value))
                {
                    this.SetXmlNodeString("@topLeftCell", value);
                }
                else
                {
                    throw new InvalidOperationException("The value must be a value cell address");
                }
            }
        }

        internal static XmlNode CreatePaneElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        {
            XmlNode? node = topNode.SelectSingleNode("d:pane", nameSpaceManager);

            if (node == null)
            {
                node = topNode.OwnerDocument.CreateElement("pane", ExcelPackage.schemaMain);
                _ = topNode.PrependChild(node);
            }

            return node;
        }
    }

    /// <summary>
    /// The selection properties for panes after a freeze or split.
    /// </summary>
    public class ExcelWorksheetPanes : XmlHelper
    {
        XmlElement _selectionNode;

        internal ExcelWorksheetPanes(XmlNamespaceManager ns, XmlNode topNode)
            : base(ns, topNode)
        {
            if (topNode.Name == "selection")
            {
                this._selectionNode = topNode as XmlElement;
            }
        }

        const string _activeCellPath = "@activeCell";

        /// <summary>
        /// Set the active cell. Must be set within the SelectedRange.
        /// </summary>
        public string ActiveCell
        {
            get
            {
                string address = this.GetXmlNodeString(_activeCellPath);

                if (address == "")
                {
                    return "A1";
                }

                return address;
            }
            set
            {
                if (this._selectionNode == null)
                {
                    this.CreateSelectionElement();
                }

                ExcelCellBase.GetRowColFromAddress(value, out int fromRow, out int fromCol, out int _, out int _);
                this.SetXmlNodeString(_activeCellPath, value);

                if (((XmlElement)this.TopNode).GetAttribute("sqref") == "")
                {
                    this.SelectedRange = ExcelCellBase.GetAddress(fromRow, fromCol);
                }
                else
                {
                    //TODO:Add fix for out of range here
                }
            }
        }

        /// <summary>
        /// The position of the pane.
        /// </summary>
        public ePanePosition Position => this.GetXmlEnumNull<ePanePosition>("@pane", ePanePosition.TopLeft).Value;

        /// <summary>
        /// 
        /// </summary>
        public int ActiveCellId { get; set; }

        private void CreateSelectionElement()
        {
            this._selectionNode = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            _ = this.TopNode.AppendChild(this._selectionNode);
            this.TopNode = this._selectionNode;
        }

        const string _selectionRangePath = "@sqref";

        /// <summary>
        /// Selected Cells. Used in combination with ActiveCell
        /// </summary>        
        public string SelectedRange
        {
            get
            {
                string address = this.GetXmlNodeString(_selectionRangePath);

                if (address == "")
                {
                    return "A1";
                }

                return address;
            }
            set
            {
                if (this._selectionNode == null)
                {
                    this.CreateSelectionElement();
                }

                ExcelCellBase.GetRowColFromAddress(value, out int fromRow, out int fromCol, out int _, out int _);
                this.SetXmlNodeString(_selectionRangePath, value);

                if (((XmlElement)this.TopNode).GetAttribute("activeCell") == "")
                {
                    this.ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
                }
                else
                {
                    //TODO:Add fix for out of range here
                }
            }
        }
    }

    private ExcelWorksheet _worksheet;

    #region ExcelWorksheetView Constructor

    /// <summary>
    /// Creates a new ExcelWorksheetView which provides access to all the view states of the worksheet.
    /// </summary>
    /// <param name="ns"></param>
    /// <param name="node"></param>
    /// <param name="xlWorksheet"></param>
    internal ExcelWorksheetView(XmlNamespaceManager ns, XmlNode node, ExcelWorksheet xlWorksheet)
        : base(ns, node)
    {
        this._worksheet = xlWorksheet;
        this.SchemaNodeOrder = new string[] { "sheetViews", "sheetView", "pane", "selection" };
        //this._paneSettings ??= new ExcelWorksheetViewPaneSettings(this.NameSpaceManager, this.TopNode);

        this.SetPaneSettings();
        this.Panes = this.LoadPanes();
    }

    private void SetPaneSettings()
    {
        XmlNode? n = this.GetNode("d:pane");

        if (n == null)
        {
            this.PaneSettings = null;
        }
        else
        {
            this.PaneSettings = new ExcelWorksheetViewPaneSettings(this.NameSpaceManager, n);
        }
    }

    #endregion

    private ExcelWorksheetPanes[] LoadPanes()
    {
        XmlNodeList nodes = this.TopNode.SelectNodes("//d:selection", this.NameSpaceManager);

        if (nodes.Count == 0)
        {
            return new ExcelWorksheetPanes[] { new ExcelWorksheetPanes(this.NameSpaceManager, this.TopNode) };
        }
        else
        {
            ExcelWorksheetPanes[] panes = new ExcelWorksheetPanes[nodes.Count];
            int i = 0;

            foreach (XmlElement elem in nodes)
            {
                panes[i++] = new ExcelWorksheetPanes(this.NameSpaceManager, elem);
            }

            return panes;
        }
    }

    #region SheetViewElement

    /// <summary>
    /// Returns a reference to the sheetView element
    /// </summary>
    protected internal XmlElement SheetViewElement => (XmlElement)this.TopNode;

    #endregion

    #region Public Methods & Properties

    /// <summary>
    /// The active cell. Single cell address.                
    /// This cell must be inside the selected range. If not, the selected range is set to the active cell address
    /// </summary>
    public string ActiveCell
    {
        get => this.Panes[this.Panes.GetUpperBound(0)].ActiveCell;
        set
        {
            ExcelAddressBase? ac = new ExcelAddressBase(value);

            if (ac.IsSingleCell == false)
            {
                throw new InvalidOperationException("ActiveCell must be a single cell.");
            }

            /*** Active cell must be inside SelectedRange ***/
            ExcelAddressBase? sd = new ExcelAddressBase(this.SelectedRange.Replace(" ", ","));
            this.Panes[this.Panes.GetUpperBound(0)].ActiveCell = value;

            if (IsActiveCellInSelection(ac, sd) == false)
            {
                this.SelectedRange = value;
            }
        }
    }

    /// <summary>
    /// The Top-Left Cell visible. Single cell address.
    /// Empty string or null is the same as A1.
    /// </summary>
    public string TopLeftCell
    {
        get => this.GetXmlNodeString("@topLeftCell");
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                this.DeleteNode("@topLeftCell");
            }
            else
            {
                if (!ExcelCellBase.IsValidCellAddress(value))
                {
                    throw new InvalidOperationException("Must be a valid cell address.");
                }

                ExcelAddressBase? ac = new ExcelAddressBase(value);

                if (ac.IsSingleCell == false)
                {
                    throw new InvalidOperationException("ActiveCell must be a single cell.");
                }

                this.SetXmlNodeString("@topLeftCell", value);
            }
        }
    }

    /// <summary>
    /// Selected Cells in the worksheet. Used in combination with ActiveCell.
    /// If the active cell is not inside the selected range, the active cell will be set to the first cell in the selected range.
    /// If the selected range has multiple adresses, these are separated with space. If the active cell is not within the first address in this list, the attribute ActiveCellId must be set (not supported, so it must be set via the XML).
    /// </summary>
    public string SelectedRange
    {
        get => this.Panes[this.Panes.GetUpperBound(0)].SelectedRange;
        set
        {
            ExcelAddressBase? ac = new ExcelAddressBase(this.ActiveCell);

            /*** Active cell must be inside SelectedRange ***/
            ExcelAddressBase? sd = new ExcelAddressBase(value.Replace(" ", ",")); //Space delimitered here, replace

            this.Panes[this.Panes.GetUpperBound(0)].SelectedRange = value;

            if (IsActiveCellInSelection(ac, sd) == false)
            {
                this.ActiveCell = new ExcelCellAddress(sd._fromRow, sd._fromCol).Address;
            }
        }
    }

    //ExcelWorksheetViewPaneSettings _paneSettings;

    /// <summary>
    /// Contains settings for the active pane
    /// </summary>
    public ExcelWorksheetViewPaneSettings PaneSettings { get; private set; }

    private static bool IsActiveCellInSelection(ExcelAddressBase ac, ExcelAddressBase sd)
    {
        ExcelAddressBase.eAddressCollition c = sd.Collide(ac);

        if (c == ExcelAddressBase.eAddressCollition.Equal || c == ExcelAddressBase.eAddressCollition.Inside)
        {
            return true;
        }
        else
        {
            if (sd.Addresses != null)
            {
                foreach (ExcelAddressBase? sds in sd.Addresses)
                {
                    c = sds.Collide(ac);

                    if (c == ExcelAddressBase.eAddressCollition.Equal || c == ExcelAddressBase.eAddressCollition.Inside)
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /// <summary>
    /// If the worksheet is selected within the workbook. NOTE: Setter clears other selected tabs.
    /// </summary>
    public bool TabSelected
    {
        get => this.GetXmlNodeBool("@tabSelected");
        set => this.SetTabSelected(value, false);
    }

    /// <summary>
    /// If the worksheet is selected within the workbook. NOTE: Setter keeps other selected tabs.
    /// </summary>
    public bool TabSelectedMulti
    {
        get => this.GetXmlNodeBool("@tabSelected");
        set => this.SetTabSelected(value, true);
    }

    /// <summary>
    /// Sets whether the worksheet is selected within the workbook.
    /// </summary>
    /// <param name="isSelected">Whether the tab is selected, defaults to true.</param>
    /// <param name="allowMultiple">Whether to allow multiple active tabs, defaults to false.</param>
    public void SetTabSelected(bool isSelected = true, bool allowMultiple = false)
    {
        if (isSelected)
        {
            this.SheetViewElement.SetAttribute("tabSelected", "1");

            if (!allowMultiple)
            {
                //    // ensure no other worksheet has its tabSelected attribute set to 1
                foreach (ExcelWorksheet sheet in this._worksheet._package.Workbook.Worksheets)
                {
                    sheet.View.TabSelected = false;
                }
            }

            XmlElement bookView = this._worksheet.Workbook.WorkbookXml.SelectSingleNode("//d:workbookView", this._worksheet.NameSpaceManager) as XmlElement;

            if (bookView != null)
            {
                bookView.SetAttribute("activeTab", this._worksheet.PositionId.ToString());
            }
        }
        else
        {
            this.SetXmlNodeString("@tabSelected", "0");
        }
    }

    /// <summary>
    /// Sets the view mode of the worksheet to pagelayout
    /// </summary>
    public bool PageLayoutView
    {
        get => this.GetXmlNodeString("@view") == "pageLayout";
        set
        {
            if (value)
            {
                this.SetXmlNodeString("@view", "pageLayout");
            }
            else
            {
                this.SheetViewElement.RemoveAttribute("view");
            }
        }
    }

    /// <summary>
    /// Sets the view mode of the worksheet to pagebreak
    /// </summary>
    public bool PageBreakView
    {
        get => this.GetXmlNodeString("@view") == "pageBreakPreview";
        set
        {
            if (value)
            {
                this.SetXmlNodeString("@view", "pageBreakPreview");
            }
            else
            {
                this.SheetViewElement.RemoveAttribute("view");
            }
        }
    }

    /// <summary>
    /// Show gridlines in the worksheet
    /// </summary>
    public bool ShowGridLines
    {
        get => this.GetXmlNodeBool("@showGridLines", true);
        set => this.SetXmlNodeString("@showGridLines", value ? "1" : "0");
    }

    /// <summary>
    /// Show the Column/Row headers (containg column letters and row numbers)
    /// </summary>
    public bool ShowHeaders
    {
        get => this.GetXmlNodeBool("@showRowColHeaders", true);
        set => this.SetXmlNodeString("@showRowColHeaders", value ? "1" : "0");
    }

    /// <summary>
    /// Window zoom magnification for current view representing percent values.
    /// </summary>
    public int ZoomScale
    {
        get => this.GetXmlNodeInt("@zoomScale");
        set
        {
            if (value < 10 || value > 400)
            {
                throw new ArgumentOutOfRangeException("Zoome scale out of range (10-400)");
            }

            this.SetXmlNodeString("@zoomScale", value.ToString());
        }
    }

    /// <summary>
    /// If the sheet is in 'right to left' display mode. Column A is on the far right and column B to the left of A. Text is also 'right to left'.
    /// </summary>
    public bool RightToLeft
    {
        get => this.GetXmlNodeBool("@rightToLeft");
        set => this.SetXmlNodeString("@rightToLeft", value == true ? "1" : "0");
    }

    internal bool WindowProtection
    {
        get => this.GetXmlNodeBool("@windowProtection", false);
        set => this.SetXmlNodeBool("@windowProtection", value, false);
    }

    /// <summary>
    /// Reference to the panes
    /// </summary>
    public ExcelWorksheetPanes[] Panes { get; internal set; }

    /// <summary>
    /// The top left pane or the top pane if the sheet is horizontaly split. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
    /// </summary>
    public ExcelWorksheetPanes TopLeftPane => this.Panes?.Where(x => x.Position == ePanePosition.TopLeft).FirstOrDefault();

    /// <summary>
    /// The top right pane. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
    /// </summary>
    public ExcelWorksheetPanes TopRightPane => this.Panes?.Where(x => x.Position == ePanePosition.TopRight).FirstOrDefault();

    /// <summary>
    /// The bottom left pane. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
    /// </summary>
    public ExcelWorksheetPanes BottomLeftPane => this.Panes?.Where(x => x.Position == ePanePosition.BottomLeft).FirstOrDefault();

    /// <summary>
    /// The bottom right pane. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
    /// </summary>
    public ExcelWorksheetPanes BottomRightPane => this.Panes?.Where(x => x.Position == ePanePosition.BottomRight).FirstOrDefault();

    string _paneNodePath = "d:pane";
    string _selectionNodePath = "d:selection";

    /// <summary>
    /// Freeze the columns/rows to left and above the cell
    /// </summary>
    /// <param name="Row"></param>
    /// <param name="Column"></param>
    public void FreezePanes(int Row, int Column)
    {
        //TODO:fix this method to handle splits as well.
        ValidateRows(Row, Column);

        if (Row == 1 && Column == 1)
        {
            this.UnFreezePanes();

            return;
        }

        bool isSplit;

        if (this.PaneSettings == null)
        {
            XmlNode? node = ExcelWorksheetViewPaneSettings.CreatePaneElement(this.NameSpaceManager, this.TopNode);
            this.PaneSettings = new ExcelWorksheetViewPaneSettings(this.NameSpaceManager, node);
            isSplit = false;
        }
        else
        {
            isSplit = this.PaneSettings.State != ePaneState.Frozen;
            this.PaneSettings.TopNode.RemoveAll();
        }

        if (Column > 1)
        {
            this.PaneSettings.XSplit = Column - 1;
        }

        if (Row > 1)
        {
            this.PaneSettings.YSplit = Row - 1;
        }

        this.PaneSettings.TopLeftCell = ExcelCellBase.GetAddress(Row, Column);
        this.PaneSettings.State = isSplit ? ePaneState.FrozenSplit : ePaneState.Frozen;

        this.CreateSelectionXml(Row - 1, Column - 1, false);
        this.Panes = this.LoadPanes();
    }

    private void CreateSelectionXml(int Row, int Column, bool isSplit)
    {
        this.RemoveSelection();

        string sqRef = this.SelectedRange,
               activeCell = this.ActiveCell;

        this.PaneSettings.ActivePanePosition = ePanePosition.BottomRight;
        XmlNode afterNode;

        if (isSplit)
        {
            //Top left node, default pane
            afterNode = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            _ = this.PaneSettings.TopNode.ParentNode.InsertAfter(afterNode, this.PaneSettings.TopNode);
        }
        else
        {
            afterNode = this.PaneSettings.TopNode;
        }

        if (Row > 0 && Column == 0)
        {
            this.PaneSettings.ActivePanePosition = ePanePosition.BottomLeft;
            XmlElement sel = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            sel.SetAttribute("pane", "bottomLeft");

            if (activeCell != "")
            {
                sel.SetAttribute("activeCell", activeCell);
            }

            if (sqRef != "")
            {
                sel.SetAttribute("sqref", sqRef);
            }

            _ = this.TopNode.InsertAfter(sel, afterNode);
        }
        else if (Column > 0 && Row == 0)
        {
            this.PaneSettings.ActivePanePosition = ePanePosition.TopRight;
            XmlElement sel = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            sel.SetAttribute("pane", "topRight");

            if (activeCell != "")
            {
                sel.SetAttribute("activeCell", activeCell);
            }

            if (sqRef != "")
            {
                sel.SetAttribute("sqref", sqRef);
            }

            _ = this.TopNode.InsertAfter(sel, afterNode);
        }
        else
        {
            XmlElement selTR = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            selTR.SetAttribute("pane", "topRight");
            string cell = ExcelCellBase.GetAddress(1, Column + 1);
            selTR.SetAttribute("activeCell", cell);
            selTR.SetAttribute("sqref", cell);
            _ = afterNode.ParentNode.InsertAfter(selTR, afterNode);

            XmlElement selBL = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            cell = ExcelCellBase.GetAddress(Row + 1, 1);
            selBL.SetAttribute("pane", "bottomLeft");
            selBL.SetAttribute("activeCell", cell);
            selBL.SetAttribute("sqref", cell);
            _ = selTR.ParentNode.InsertAfter(selBL, selTR);

            XmlElement selBR = this.TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
            selBR.SetAttribute("pane", "bottomRight");

            if (activeCell != "")
            {
                selBR.SetAttribute("activeCell", activeCell);
            }

            if (sqRef != "")
            {
                selBR.SetAttribute("sqref", sqRef);
            }

            _ = selBL.ParentNode.InsertAfter(selBR, selBL);
        }
    }

    private static void ValidateRows(int Row, int Column)
    {
        if (Row < 0 || Row > ExcelPackage.MaxRows - 1)
        {
            throw new ArgumentOutOfRangeException($"Row must not be negative or exceed {ExcelPackage.MaxRows - 1}");
        }

        if (Column < 0 || Column > ExcelPackage.MaxColumns - 1)
        {
            throw new ArgumentOutOfRangeException($"Column must not be negative or exceed {ExcelPackage.MaxColumns - 1}");
        }
    }

    /// <summary>
    /// Split panes at the position in pixels from the top-left corner.
    /// </summary>
    /// <param name="pixelsY">Vertical pixels</param>
    /// <param name="pixelsX">Horizontal pixels</param>
    public void SplitPanesPixels(int pixelsY, int pixelsX)
    {
        if (pixelsY <= 0 && pixelsX <= 0) //Both row and column is zero, remove the panes.
        {
            this.UnFreezePanes();

            return;
        }

        this.SetPaneSetting();

        ExcelCellAddress? c = this.GetTopLeftCell();

        if (pixelsX > 0)
        {
            ExcelStyles? styles = this._worksheet.Workbook.Styles;
            int normalStyleIx = styles.GetNormalStyleIndex();
            ExcelFont? nf = styles.NamedStyles[normalStyleIx < 0 ? 0 : normalStyleIx].Style.Font;
            double defaultWidth = Convert.ToDouble(FontSize.GetWidthPixels(nf.Name, nf.Size));
            int widthCharRH = c.Row < 1000 ? 3 : c.Row.ToString(CultureInfo.InvariantCulture).Length;
            int margin = 5;
            this.PaneSettings.XSplit = (Convert.ToDouble(pixelsX) + (defaultWidth * widthCharRH) + margin) * 15D;
        }

        if (pixelsY > 0)
        {
            this.PaneSettings.YSplit = (pixelsY + (this._worksheet.DefaultRowHeight / 0.75)) * 15D;
        }

        this.CreateSelectionXml(pixelsY == 0 ? 0 : 1, pixelsX == 0 ? 0 : 1, true);
        this.Panes = this.LoadPanes();

        if (pixelsX > 0 && pixelsY > 0)
        {
            ExcelCellAddress? a = new ExcelCellAddress(string.IsNullOrEmpty(this.TopLeftCell) ? "A1" : this.TopLeftCell);
            this.PaneSettings.TopLeftCell = ExcelCellBase.GetAddress(a.Row, a.Column);
        }
    }

    /// <summary>
    /// Split the window at the supplied row/column. 
    /// The split is performed using the current width/height of the visible rows and columns, so any changes to column width or row heights after the split will not effect the split position.
    /// To remove split call this method with zero as value of both paramerters or use <seealso cref="UnFreezePanes"/>
    /// </summary>
    /// <param name="rowsTop">Splits the panes at the coordinate after this visible row. Zero mean no split on row level</param>
    /// <param name="columnsLeft">Splits the panes at the coordinate after this visible column. Zero means no split on column level.</param>
    public void SplitPanes(int rowsTop, int columnsLeft)
    {
        ValidateRows(rowsTop, columnsLeft);

        if (rowsTop == 0 && columnsLeft == 0) //Both row and column is zero, remove the panes.
        {
            this.UnFreezePanes();

            return;
        }

        this.SetPaneSetting();

        ExcelCellAddress? c = this.GetTopLeftCell();

        if (columnsLeft > 0)
        {
            ExcelStyles? styles = this._worksheet.Workbook.Styles;
            int normalStyleIx = styles.GetNormalStyleIndex();
            ExcelFont? nf = styles.NamedStyles[normalStyleIx < 0 ? 0 : normalStyleIx].Style.Font;
            decimal defaultWidth = FontSize.GetWidthPixels(nf.Name, nf.Size);
            int widthCharRH = c.Row < 1000 ? 3 : c.Row.ToString(CultureInfo.InvariantCulture).Length;
            int margin = 5;
            this.PaneSettings.XSplit = Convert.ToDouble(this.GetVisibleColumnWidth(c.Column - 1, columnsLeft) + (defaultWidth * widthCharRH) + margin) * 15D;
        }

        if (rowsTop > 0)
        {
            this.PaneSettings.YSplit = (Convert.ToDouble(this.GetVisibleRowWidth(c.Row, rowsTop)) + (this._worksheet.DefaultRowHeight / 0.75)) * 15D;
        }

        this.CreateSelectionXml(rowsTop, columnsLeft, true);
        this.Panes = this.LoadPanes();

        ExcelCellAddress? a = new ExcelCellAddress(string.IsNullOrEmpty(this.TopLeftCell) ? "A1" : this.TopLeftCell);
        this.PaneSettings.TopLeftCell = ExcelCellBase.GetAddress(a.Row + rowsTop, a.Column + columnsLeft);
    }

    private void SetPaneSetting()
    {
        if (this.PaneSettings == null)
        {
            XmlNode? node = ExcelWorksheetViewPaneSettings.CreatePaneElement(this.NameSpaceManager, this.TopNode);
            this.PaneSettings = new ExcelWorksheetViewPaneSettings(this.NameSpaceManager, node);
        }
        else
        {
            this.PaneSettings.State = ePaneState.Split;
        }
    }

    private ExcelCellAddress GetTopLeftCell()
    {
        if (string.IsNullOrEmpty(this.TopLeftCell))
        {
            if (string.IsNullOrEmpty(this.PaneSettings?.TopLeftCell))
            {
                return new ExcelCellAddress();
            }
            else
            {
                return new ExcelCellAddress(this.PaneSettings.TopLeftCell);
            }
        }
        else
        {
            return new ExcelCellAddress(this.TopLeftCell);
        }
    }

    private decimal GetVisibleColumnWidth(int topCol, int cols)
    {
        decimal mdw = this._worksheet.Workbook.MaxFontWidth;
        decimal width = 0;

        for (int c = 0; c < cols; c++)
        {
            width += this._worksheet.GetColumnWidthPixels(topCol + c, mdw);
        }

        return width;
    }

    private decimal GetVisibleRowWidth(int leftRow, int rows)
    {
        decimal height = 0;

        for (int r = 0; r < rows; r++)
        {
            height += Convert.ToDecimal(this._worksheet.GetRowHeight(leftRow + r)) / 0.75M;
        }

        return height;
    }

    private void RemoveSelection()
    {
        //Find selection nodes and remove them            
        XmlNodeList selections = this.TopNode.SelectNodes(this._selectionNodePath, this.NameSpaceManager);

        foreach (XmlNode sel in selections)
        {
            _ = sel.ParentNode.RemoveChild(sel);
        }
    }

    /// <summary>
    /// Unlock all rows and columns to scroll freely
    /// </summary>
    public void UnFreezePanes()
    {
        string sqRef = this.SelectedRange,
               activeCell = this.ActiveCell;

        XmlElement paneNode = this.TopNode.SelectSingleNode(this._paneNodePath, this.NameSpaceManager) as XmlElement;

        if (paneNode != null)
        {
            _ = paneNode.ParentNode.RemoveChild(paneNode);
        }

        this.RemoveSelection();

        this.PaneSettings = null;
        this.Panes = new ExcelWorksheetPanes[] { new ExcelWorksheetPanes(this.NameSpaceManager, this.TopNode) };

        this.SelectedRange = sqRef;
        this.ActiveCell = activeCell;
    }

    #endregion
}