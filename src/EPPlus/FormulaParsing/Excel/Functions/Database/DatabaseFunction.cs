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
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

internal abstract class DatabaseFunction : ExcelFunction
{
    protected RowMatcher RowMatcher { get; private set; }

    public DatabaseFunction()
        : this(new RowMatcher())
    {
    }

    public DatabaseFunction(RowMatcher rowMatcher) => this.RowMatcher = rowMatcher;

    protected IEnumerable<double> GetMatchingValues(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        string? dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;

        //var field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
        object? field = arguments.ElementAt(1).Value;
        string? criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;

        ExcelDatabase? db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
        ExcelDatabaseCriteria? criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);
        List<double>? values = new List<double>();

        while (db.HasMoreRows)
        {
            ExcelDatabaseRow? dataRow = db.Read();

            if (!this.RowMatcher.IsMatch(dataRow, criteria))
            {
                continue;
            }

            object? candidate = ConvertUtil.IsNumericOrDate(field)
                                    ? dataRow[(int)ConvertUtil.GetValueDouble(field)]
                                    : dataRow[field.ToString().ToLower(CultureInfo.InvariantCulture)];

            if (ConvertUtil.IsNumericOrDate(candidate))
            {
                values.Add(ConvertUtil.GetValueDouble(candidate));
            }
        }

        return values;
    }
}