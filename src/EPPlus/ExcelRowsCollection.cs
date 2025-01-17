﻿/*************************************************************************************************
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

using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml;

/// <summary>
/// A collection of rows in a worksheet.
/// </summary>
public class ExcelRowsCollection : ExcelRangeRow
{
    ExcelWorksheet _worksheet;

    internal ExcelRowsCollection(ExcelWorksheet worksheet)
        : base(worksheet, 1, ExcelPackage.MaxRows) =>
        this._worksheet = worksheet;

    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="row">The row index</param>
    /// <returns>The <see cref="ExcelRangeRow"/></returns>
    public ExcelRangeRow this[int row] => new(this._worksheet, row, row);

    /// <summary>
    /// Indexer for the collection
    /// </summary>
    /// <param name="fromRow">The row index from which collection should start</param>
    /// <param name="toRow">index from which collection should end</param>
    /// <returns>The <see cref="ExcelRangeRow"/></returns>
    public ExcelRangeRow this[int fromRow, int toRow] => new(this._worksheet, fromRow, toRow);
}