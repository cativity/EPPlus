/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/18/2021         EPPlus Software AB       Improved handling of hidden cells for SUBTOTAL and AGGREGATE.
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using OfficeOpenXml.Table;

namespace OfficeOpenXml.FormulaParsing;

/// <summary>
/// This class contains information of the usage of Filters on the worksheets of a workbook.
/// One area where this information is needed is when running the SUBTOTAL function. If
/// there is an active filter on the worksheet hidden cells should be ignored even if SUBTOTAL
/// is called with a single digit func num.
/// </summary>
internal class FilterInfo
{
    public FilterInfo(ExcelWorkbook workbook)
    {
        this._workbook = workbook;
        this.Initialize();
    }

    private readonly ExcelWorkbook _workbook;
    private readonly HashSet<string> _worksheetFilters = new HashSet<string>();

    private void Initialize()
    {
        foreach (ExcelWorksheet? worksheet in this._workbook.Worksheets)
        {
            if (worksheet.IsChartSheet)
            {
                continue;
            }

            if (worksheet.AutoFilter != null && worksheet.AutoFilter.Columns != null && worksheet.AutoFilter.Columns.Count > 0)
            {
                this._worksheetFilters.Add(worksheet.Name);

                continue;
            }

            foreach (ExcelTable? table in worksheet.Tables)
            {
                if (table.AutoFilter != null && table.AutoFilter.Columns != null && table.AutoFilter.Columns.Count > 0)
                {
                    if (!this._worksheetFilters.Contains(worksheet.Name))
                    {
                        this._worksheetFilters.Add(worksheet.Name);

                        continue;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Returns true if there is an Autofilter with at least one column on the requested worksheet.
    /// </summary>
    /// <param name="worksheetName"></param>
    /// <returns></returns>
    public bool WorksheetHasFilter(string worksheetName)
    {
        return this._worksheetFilters.Contains(worksheetName);
    }
}