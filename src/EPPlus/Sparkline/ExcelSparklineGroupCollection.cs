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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sparkline;

/// <summary>
/// A collection of sparkline groups
/// </summary>
public class ExcelSparklineGroupCollection : IEnumerable<ExcelSparklineGroup>
{
    ExcelWorksheet _ws;
    List<ExcelSparklineGroup> _lst;

    internal ExcelSparklineGroupCollection(ExcelWorksheet ws)
    {
        this._ws = ws;
        this._lst = new List<ExcelSparklineGroup>();
        this.LoadSparklines();
    }

    const string _extPath = "/d:worksheet/d:extLst/d:ext";
    const string _searchPath = "[@uri='{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']";
    const string _topSearchPath = _extPath + _searchPath + "/x14:sparklineGroups";
    const string _topPath = _extPath + "/x14:sparklineGroups";

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count => this._lst.Count;

    /// <summary>
    /// Adds a new sparklinegroup to the collection
    /// </summary>
    /// <param name="type">Type of sparkline</param>
    /// <param name="locationRange">The location of the sparkline group. The range must have one row or column and must match the number of rows/columns in the datarange</param>
    /// <param name="dataRange">The data for the sparkline group</param>
    /// <returns></returns>
    public ExcelSparklineGroup Add(eSparklineType type, ExcelAddressBase locationRange, ExcelAddressBase dataRange)
    {
        if (locationRange.Rows == 1)
        {
            if (locationRange.Columns == dataRange.Rows)
            {
                return this.AddGroup(type, locationRange, dataRange, true);
            }
            else if (locationRange.Columns == dataRange.Columns)
            {
                return this.AddGroup(type, locationRange, dataRange, false);
            }
            else
            {
                throw new ArgumentException("dataRange is not valid. dataRange columns or rows must match number of rows in locationRange");
            }
        }
        else if (locationRange.Columns == 1)
        {
            if (locationRange.Rows == dataRange.Columns)
            {
                return this.AddGroup(type, locationRange, dataRange, false);
            }
            else if (locationRange.Rows == dataRange.Rows)
            {
                return this.AddGroup(type, locationRange, dataRange, true);
            }
            else
            {
                throw new ArgumentException("dataRange is not valid. dataRange columns or rows must match number of columns in locationRange");
            }
        }
        else
        {
            throw new ArgumentException("locationRange is not valid. Range must be one Column or Row only");
        }
    }

    private ExcelSparklineGroup AddGroup(eSparklineType type, ExcelAddressBase locationRange, ExcelAddressBase dataRange, bool isRows)
    {
        ExcelSparklineGroup? group = this.NewSparklineGroup();
        group.Type = type;
        int row = locationRange._fromRow;
        int col = locationRange._fromCol;

        int drFromRow = dataRange._fromRow;
        int drFromCol = dataRange._fromCol;
        int drToRow = isRows ? dataRange._fromRow : dataRange._toRow;
        int drToCol = isRows ? dataRange._toCol : dataRange._fromCol;

        int cells = (locationRange._fromRow == locationRange._toRow
                         ? locationRange._toCol - locationRange._fromCol
                         : locationRange._toRow - locationRange._fromRow)
                    + 1;

        int cell = 0;
        string? wsName = dataRange?.WorkSheetName ?? this._ws.Name;

        while (cell < cells)
        {
            ExcelCellAddress? f = new ExcelCellAddress(row, col);
            ExcelAddressBase? sqref = new ExcelAddressBase(wsName, drFromRow, drFromCol, drToRow, drToCol);
            group.Sparklines.Add(f, wsName, sqref);
            cell++;

            if (locationRange._fromRow == locationRange._toRow)
            {
                col++;
            }
            else
            {
                row++;
            }

            if (isRows)
            {
                drFromRow++;
                drToRow++;
            }
            else
            {
                drFromCol++;
                drToCol++;
            }
        }

        group.ColorSeries.Rgb = "FF376092";
        group.ColorNegative.Rgb = "FFD00000";
        group.ColorMarkers.Rgb = "FFD00000";
        group.ColorAxis.Rgb = "FF000000";
        group.ColorFirst.Rgb = "FFD00000";
        group.ColorLast.Rgb = "FFD00000";
        group.ColorHigh.Rgb = "FFD00000";
        group.ColorLow.Rgb = "FFD00000";
        this._lst.Add(group);

        return group;
    }

    private ExcelSparklineGroup NewSparklineGroup()
    {
        XmlHelperInstance? xh =
            new XmlHelperInstance(this._ws.NameSpaceManager, this._ws.WorksheetXml); //SelectSingleNode("/d:worksheet", _ws.NameSpaceManager)

        if (!xh.ExistsNode(_extPath + _searchPath))
        {
            XmlNode? ext = xh.CreateNode(_extPath, false, true);

            if (ext.Attributes["uri"] == null)
            {
                ((XmlElement)ext).SetAttribute("uri",
                                               "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"); //Guid URI for sparklines... https://msdn.microsoft.com/en-us/library/dd905242(v=office.12).aspx
            }
        }

        XmlNode? parent = xh.CreateNode(_topSearchPath);

        XmlElement? topNode = this._ws.WorksheetXml.CreateElement("x14", "sparklineGroup", ExcelPackage.schemaMainX14);
        topNode.SetAttribute("xmlns:xm", ExcelPackage.schemaMainXm);
        _ = topNode.SetAttribute("uid", ExcelPackage.schemaXr2, $"{{{Guid.NewGuid().ToString()}}}");
        _ = parent.AppendChild(topNode);
        _ = topNode.AppendChild(topNode.OwnerDocument.CreateElement("x14", "sparklines", ExcelPackage.schemaMainX14));

        return new ExcelSparklineGroup(this._ws.NameSpaceManager, topNode, this._ws);
    }

    private void LoadSparklines()
    {
        XmlNodeList? grps = this._ws.WorksheetXml.SelectNodes(_topPath + "/x14:sparklineGroup", this._ws.NameSpaceManager);

        foreach (XmlElement grp in grps)
        {
            this._lst.Add(new ExcelSparklineGroup(this._ws.NameSpaceManager, grp, this._ws));
        }
    }

    /// <summary>
    /// Returns the sparklinegroup at the specified position.  
    /// </summary>
    /// <param name="index">The position of the Sparklinegroup. 0-base</param>
    /// <returns></returns>
    public ExcelSparklineGroup this[int index] => this._lst[index];

    /// <summary>
    /// The enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelSparklineGroup> GetEnumerator() => this._lst.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this._lst.GetEnumerator();

    /// <summary>
    /// Removes the sparkline.
    /// </summary>
    /// <param name="index">The index of the item to be removed</param>
    public void RemoveAt(int index) => this.Remove(this._lst[index]);

    /// <summary>
    /// Removes the sparkline.
    /// </summary>
    /// <param name="sparklineGroup">The sparklinegroup to be removed</param>
    public void Remove(ExcelSparklineGroup sparklineGroup)
    {
        _ = sparklineGroup.TopNode.ParentNode.RemoveChild(sparklineGroup.TopNode);
        _ = this._lst.Remove(sparklineGroup);
    }
}